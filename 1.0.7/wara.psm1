<#
.SYNOPSIS
    Starts the WARA Collector process.

.DESCRIPTION
    The Start-WARACollector function initiates the WARA Collector process, which collects and processes data based on the specified parameters. It supports multiple parameter sets, including Default, Specialized, and ConfigFileSet.

.PARAMETER SAP
    Switch to enable SAP workload processing.

.PARAMETER AVD
    Switch to enable AVD workload processing.

.PARAMETER AVS
    Switch to enable AVS workload processing.

.PARAMETER HPC
    Switch to enable HPC workload processing.

.PARAMETER AI_GPT_RAG
    Switch to enable Artificial Intelligence (GPT-RAG) workload processing.

.PARAMETER ORACLE
    Switch to enable Oracle workload processing.

.PARAMETER PassThru
    Switch to enable the PassThru parameter. PassThru returns the output object.

.PARAMETER SubscriptionIds
    Array of subscription IDs to include in the process. Validated using Test-WAFSubscriptionId.

.PARAMETER ResourceGroups
    Array of resource groups to include in the process. Validated using Test-WAFResourceGroupId.

.PARAMETER TenantID
    The tenant ID to use for the process. This parameter is mandatory and validated using Test-WAFIsGuid.

.PARAMETER Tags
    Array of tags to include in the process. Validated using Test-WAFTagPattern.

.PARAMETER AzureEnvironment
    Specifies the Azure environment to connect to. If omitted, the module inherits the environment of the current Az context (Get-AzContext), so it works against any registered sovereign cloud (e.g. USNAT, USSec, AzureUSGovernment) as well as the commercial cloud. Sovereign environments must be registered first with Add-AzEnvironment.

.PARAMETER ConfigFile
    Path to the configuration file. This parameter is mandatory for the ConfigFileSet parameter set and validated using Test-Path.

.PARAMETER RecommendationDataUri
    Local file path (preferred for disconnected/sovereign clouds) or http(s) URI for the recommendation data. Defaults to the bundled copy under the module's offline-data folder.

.PARAMETER RecommendationResourceTypesUri
    Local file path (preferred for disconnected/sovereign clouds) or http(s) URI for the recommendation resource types CSV. Defaults to the bundled copy under the module's offline-data folder.

.PARAMETER SkipVersionCheck
    Skips the PowerShell Gallery version check. Useful in disconnected environments where the Gallery is unreachable.

.PARAMETER UseImplicitRunbookSelectors
    Switch to enable the use of implicit runbook selectors.

.PARAMETER RunbookFile
    Path to the runbook file. Validated using Test-Path.

.EXAMPLE
    Start-WARACollector -TenantID "00000000-0000-0000-0000-000000000000" -SubscriptionIds "/subscriptions/00000000-0000-0000-0000-000000000000"

.EXAMPLE
    Start-WARACollector -ConfigFile "C:\path\to\config.txt"

.EXAMPLE
    Start-WARACollector -TenantID "00000000-0000-0000-0000-000000000000" -SubscriptionIds "/subscriptions/00000000-0000-0000-0000-000000000000" -ResourceGroups "/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/RG-001" -Tags "Env||Environment!~Dev||QA" -AVD -SAP -HPC -ORACLE

.EXAMPLE
    Start-WARACollector -ConfigFile "C:\path\to\config.txt" -SAP -AVD

.NOTES
    Author: Kyle Poineal
    Date: 12/11/2024
#>
function Start-WARACollector {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $SAP,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $AVD,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $AVS,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $HPC,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $AI_GPT_RAG,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $ORACLE,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $PassThru,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [ValidateScript({ Test-WAFSubscriptionId $_ })]
        [string[]] $SubscriptionIds,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [ValidateScript({ Test-WAFResourceGroupId $_ })]
        [string[]] $ResourceGroups,

        [Parameter(Mandatory = $true, ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [ValidateScript({ Test-WAFIsGuid $_ })]
        [GUID] $TenantID,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [ValidateScript({ Test-WAFTagPattern $_ })]
        [string[]] $Tags,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [string] $AzureEnvironment = '',

        [Parameter(ParameterSetName = 'ConfigFileSet', Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string] $ConfigFile,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [ValidateScript({ ($_ -match '^https?://') -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
        [string] $RecommendationDataUri = (Join-Path $PSScriptRoot 'offline-data/recommendations.json'),

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [ValidateScript({ ($_ -match '^https?://') -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
        [string] $RecommendationResourceTypesUri = (Join-Path $PSScriptRoot 'offline-data/WARAinScopeResTypes.csv'),

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Specialized')]
        [Parameter(ParameterSetName = 'ConfigFileSet')]
        [switch] $SkipVersionCheck
    )

    # Check for module updates. In disconnected / sovereign environments (e.g. USNAT) the PowerShell
    # Gallery is not reachable, so this is best-effort only: it never throws, and can be bypassed
    # entirely with -SkipVersionCheck.
    <#
    if (-not $SkipVersionCheck) {
        try {
            Write-Host 'Checking Version..' -ForegroundColor Cyan
            $LocalVersion = (Get-Module -Name $MyInvocation.MyCommand.ModuleName).Version
            $GalleryVersion = (Find-Module -Name $MyInvocation.MyCommand.ModuleName -ErrorAction Stop).Version

            if ($LocalVersion -lt $GalleryVersion) {
                Write-Warning "A newer version ($GalleryVersion) of $($MyInvocation.MyCommand.ModuleName) is available (installed: $LocalVersion). Consider updating with 'Update-Module -Name $($MyInvocation.MyCommand.ModuleName)'."
            }
        }
        catch {
            Write-Verbose "Version check skipped - module repository not reachable: $($_.Exception.Message)"
        }
    }
    #>

    # Start the stopwatch to time the script
    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Create a hashtable of the parameters passed to the script
    $scriptParams = foreach ($param in $PSBoundParameters.GetEnumerator()) {
        Write-Debug "Parameter: $($param.Key) Value: $($param.Value)"
        [PSCustomObject]@{
            $param.key = $param.value
        }
    }

    Write-Debug 'Debugging mode is enabled'
    Write-Progress -Activity 'WARA Collector' -Status 'Starting WARA Collector' -PercentComplete 0 -Id 1

    # Determine which parameter set is active
    switch ($PSCmdlet.ParameterSetName) {
        'ConfigFileSet' {
            Write-Debug 'Using ConfigFileSet parameter set'
            Write-Debug "ConfigFile: $ConfigFile"
            Write-Debug 'Importing ConfigFile data'
            $ConfigData = Import-WAFConfigFileData -ConfigFile $ConfigFile
            Write-Debug 'Testing TenantId, SubscriptionIds, ResourceGroups, and Tags'
            $ConfigData.TenantId = ([guid][string]$ConfigData.TenantId).Guid
            $null = Test-WAFIsGuid -StringGuid $ConfigData.TenantId
            $null = if ($ConfigData.SubscriptionIds) { Test-WAFSubscriptionId -InputValue $ConfigData.SubscriptionIds }
            $null = if ($ConfigData.ResourceGroups) { Test-WAFResourceGroupId -InputValue $ConfigData.ResourceGroups }
            $null = if ($ConfigData.Tags) { Test-WAFTagPattern -InputValue $ConfigData.Tags }
        }
        'Default' {
            Write-Debug 'Using Default parameter set'
            Write-Debug "Parameter set values: $($PSBoundParameters.Keys)"

            if ($PSBoundParameters.keys.contains('SubscriptionIds') -or $PSBoundParameters.keys.contains('ResourceGroups')) {
                Write-Debug 'We contain the parameters.'
            }
            else {
                Write-Debug 'We do not contain the parameters.'
                throw 'The parameter SubscriptionIds or ResourceGroups is required when using the Default parameter set.'
            }
        }
    }


    #Use Null Coalescing to set the values of parameters.
    Write-Progress -Activity 'WARA Collector' -Status 'Setting Scope' -PercentComplete 1 -Id 1
    $Scope_TenantId = $ConfigData.TenantId ?? $TenantID ?? (throw 'Tenant ID is required.')
    $Scope_SubscriptionIds = $ConfigData.SubscriptionIds ?? $SubscriptionIds ?? @()
    $Scope_ResourceGroups = $ConfigData.ResourceGroups ?? $ResourceGroups ?? @()
    $Scope_Tags = $ConfigData.Tags ?? $Tags ?? @()

    $Scope_TenantId = ([guid][string]$Scope_TenantId).Guid

    Write-Progress -Activity 'WARA Collector' -Status 'Setting Scope' -PercentComplete 3 -Id 1
    $Scope_SubscriptionIds = Repair-WAFSubscriptionId -SubscriptionIds $Scope_SubscriptionIds

    Write-Debug "Tenant ID: $Scope_TenantId"
    Write-Debug "Subscription IDs: $Scope_SubscriptionIds"
    Write-Debug "Resource Groups: $Scope_ResourceGroups"
    Write-Debug "Tags: $Scope_Tags"

    # Create Specialized Workloads array
    $SpecializedWorkloads = @()

    # Check if any of the specialized workloads are enabled and add them to $SpecializedWorkloads
    if ($SAP) {
        Write-Debug 'SAP switch is enabled'
        $SpecializedWorkloads += 'SAP'
    }
    if ($AVD) {
        Write-Debug 'AVD switch is enabled'
        $SpecializedWorkloads += 'AVD'
    }
    if ($AVS) {
        Write-Debug 'AVS switch is enabled'
        $SpecializedWorkloads += 'AVS'
    }
    if ($HPC) {
        Write-Debug 'HPC switch is enabled'
        $SpecializedWorkloads += 'HPC'
    }
    if ($AI_GPT_RAG) {
        Write-Debug 'AI_GPT_RAG switch is enabled'
        $SpecializedWorkloads += 'AI-GPT-RAG'
    }
    if ($ORACLE) {
        Write-Debug 'ORACLE switch is enabled'
        $SpecializedWorkloads += 'ORACLE'
    }

    if ($SpecializedWorkloads) {
        Write-Debug "Specialized Workloads: $SpecializedWorkloads"
    }

    #Import Recommendation Object from WARA-Build GitHub Pages Site
    Write-Progress -Activity 'WARA Collector' -Status 'Importing APRL Recommendation Object from GitHub' -PercentComplete 5 -Id 1
    Write-Debug 'Importing APRL Recommendation Object'
    $RecommendationObject = Get-WAFDataSource -Source $RecommendationDataUri
    if ($RecommendationObject -is [string]) { $RecommendationObject = $RecommendationObject | ConvertFrom-Json }
    Write-Debug "Count of APRL Recommendation Object: $($RecommendationObject.count)"

    #Create Recommendation Object HashTable for faster lookup
    Write-Debug 'Creating Recommendation Object HashTable for faster lookup'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating Recommendation Object HashTable' -PercentComplete 8 -Id 1
    $RecommendationObjectHash = @{}
    $RecommendationObject.ForEach({ $RecommendationObjectHash[$_.aprlGuid] = $_ })
    Write-Debug "Count of Recommendation Object Hashtable: $($RecommendationObjectHash.count)"

    #Import WARA InScope Resource Types CSV from APRL
    Write-Debug 'Importing WARA InScope Resource Types CSV'
    Write-Progress -Activity 'WARA Collector' -Status 'Importing WARA InScope Resource Types CSV' -PercentComplete 11 -Id 1
    $RecommendationResourceTypes = Get-WAFDataSource -Source $RecommendationResourceTypesUri
    $RecommendationResourceTypes = $RecommendationResourceTypes | ConvertFrom-Csv | Where-Object { $_.WARAinScope -eq 'yes' }
    Write-Debug "Count of WARA InScope Resource Types: $($RecommendationResourceTypes.count)"

    #Add Specialized Workloads to WARA InScope Resource Types
    Write-Debug 'Adding Specialized Workloads to WARA InScope Resource Types'
    Write-Progress -Activity 'WARA Collector' -Status 'Adding Specialized Workloads to WARA InScope Resource Types' -PercentComplete 14 -Id 1
    $RecommendationResourceTypes += $SpecializedWorkloads
    Write-Debug "Count of WARA InScope Resource Types with Specialized Workloads: $($RecommendationResourceTypes.count)"

    #Create TypesNotInAPRLOrAdvisor Object from WARA InScope Resource Types
    Write-Debug 'Creating TypesNotInAPRLOrAdvisor Object from WARA InScope Resource Types'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating TypesNotInAPRLOrAdvisor Object' -PercentComplete 17 -Id 1
    $TypesNotInAPRLOrAdvisor = ($RecommendationResourceTypes | Where-Object { $_.InAprlAndOrAdvisor -eq "No" }).ResourceType
    Write-Debug "Count of TypesNotInAPRLOrAdvisor: $($TypesNotInAPRLOrAdvisor.count)"

    #Connect to Azure
    Write-Debug 'Connecting to Azure if not connected.'
    Write-Progress -Activity 'WARA Collector' -Status 'Validating connection to Azure' -PercentComplete 20 -Id 1
    Connect-WAFAzure -TenantID $Scope_TenantId -AzureEnvironment $AzureEnvironment

    # Resolve the effective environment name and ARM endpoint from the (now-established) context so the
    # rest of the collector targets the correct sovereign endpoints regardless of which cloud we are in.
    $AzureEnvironment = (Get-AzContext).Environment.Name
    $BaseURL = (Get-AzContext).Environment.ResourceManagerUrl
    Write-Debug "Using Azure Environment: $AzureEnvironment  ResourceManagerUrl: $BaseURL"

    # Track any optional data sources (Advisor / Resource Health / Support) that cannot be reached in the
    # target cloud, so collection degrades gracefully and records the gap in the output instead of
    # aborting the whole run (R2).
    $collectionErrors = [System.Collections.Generic.List[object]]::new()

    #Get Implicit Subscription Ids from Scope
    Write-Debug 'Getting Implicit Subscription Ids from Scope'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Implicit Subscription Ids' -PercentComplete 23 -Id 1
    $Scope_ImplicitSubscriptionIds = Get-WAFImplicitSubscriptionId -SubscriptionFilters $Scope_SubscriptionIds -ResourceGroupFilters $Scope_ResourceGroups
    Write-Debug "Implicit Subscription Ids: $Scope_ImplicitSubscriptionIds"

    #Get all resources from the Implicit Subscription ID scope - We use this later to add type, location, subscriptionid, resourcegroup to the impactedResourceObj objects
    Write-Debug 'Getting all resources from the Implicit Subscription ID scope'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting All Resources' -PercentComplete 26 -Id 1
    $AllResources = Invoke-WAFQuery -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
    Write-Debug "Count of Resources: $($AllResources.count)"

    #Create HashTable of all resources for faster lookup
    Write-Debug 'Creating HashTable of all resources for faster lookup'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating All Resources HashTable' -PercentComplete 29 -Id 1
    $AllResourcesHash = @{}
    Foreach ($resource in $AllResources) {
        $AllResourcesHash[$resource.id] = $resource
    }
    #$AllResources.ForEach({ $AllResourcesHash[$_.id] = $_ })
    Write-Debug "All Resources Hash: $($AllResourcesHash.count)"

    #Filter all resources by subscription, resourcegroup, and resource scope
    Write-Debug 'Filtering all resources by subscription, resourcegroup, and resource scope'
    Write-Progress -Activity 'WARA Collector' -Status 'Filtering All Resources' -PercentComplete 32 -Id 1
    $Scope_AllResources = Get-WAFFilteredResourceList -UnfilteredResources $AllResources -SubscriptionFilters $Scope_SubscriptionIds -ResourceGroupFilters $Scope_ResourceGroups
    Write-Debug "Count of filtered Resources: $($Scope_AllResources.count)"

    #Create Resource Inventory object
    Write-Debug 'Creating Resource Inventory object'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating Resource Inventory' -PercentComplete 33 -Id 1
    $ResourceInventory = $Scope_AllResources
    Write-Debug "Count of Resource Inventory: $($ResourceInventory.count)"


    #Filter all resources by InScope Resource Types - We do this because we need to be able to compare resource ids to generate the generic recommendations(Resource types that have no recommendations or are not in advisor but also need to be validated)
    Write-Debug 'Filtering all resources by WARA InScope Resource Types'
    Write-Progress -Activity 'WARA Collector' -Status 'Filtering All Resources by WARA InScope Resource Types' -PercentComplete 35 -Id 1
    $Scope_AllResources = Get-WAFResourcesByList -ObjectList $Scope_AllResources -FilterList $RecommendationResourceTypes.ResourceType -KeyColumn 'type'
    Write-Debug "Count of filtered by type Resources: $($Scope_AllResources.count)"


    #Get all APRL recommendations from the Implicit Subscription ID scope
    Write-Debug 'Getting all APRL recommendations from the Implicit Subscription ID scope'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting APRL Recommendations' -PercentComplete 38 -Id 1
    $Recommendations = Invoke-WAFQueryLoop -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '') -RecommendationObject $RecommendationObject -AddedTypes $SpecializedWorkloads -ProgressId 2
    Write-Debug "Count of Recommendations: $($Recommendations.count)"

    #Filter resource recommendation objects by subscription, resourcegroup, and resource scope
    Write-Debug 'Filtering APRL recommendation objects by subscription, resourcegroup, and resource scope'
    Write-Progress -Activity 'WARA Collector' -Status 'Filtering APRL Recommendations' -PercentComplete 41 -Id 1
    $Filter_Recommendations = Get-WAFFilteredResourceList -UnfilteredResources $Recommendations -SubscriptionFilters $Scope_SubscriptionIds -ResourceGroupFilters $Scope_ResourceGroups
    Write-Debug "Count of APRL recommendation objects: $($Filter_Recommendations.count)"

    #Create impactedResourceObj objects from the recommendations
    Write-Debug 'Creating impactedResourceObj objects from the recommendations'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating Impacted Resource Objects' -PercentComplete 44 -Id 1
    $impactedResourceObj = Build-ImpactedResourceObj -ImpactedResource $Filter_Recommendations -AllResources $AllResourcesHash -RecommendationObject $RecommendationObjectHash
    Write-Debug "Count of impactedResourceObj objects: $($impactedResourceObj.count)"

    #Create list of validationResourceIds from the impactedResourceObj objects
    Write-Debug 'Creating hashtable of validationResources from the impactedResourceObj objects'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating Validation Resources' -PercentComplete 47 -Id 1
    $validationResources = @{}
    foreach ($obj in $impactedResourceObj | Select-Object id, name, type, location, subscriptionid, resourcegroup, checkname, selector) {
        $key = "$($obj.id)"
        if (-not $validationResources.ContainsKey($key)) {
            $validationResources[$key] = $obj
        }
    }
    Write-Debug "Count of validationResourceIds: $($validationResources.count)"

    #Add In Scope resources to validationResources HashTable
    #By adding the $Scope_AllResources to the validationResources HashTable, we can ensure that we have all resources in the scope that need to be validated.
    #Adding the resources AFTER the first loop ensures that we do not add resources that are already in the impactedResourceObj objects.
    #This means we do not have to worry about overwriting the objects.
    Write-Debug 'Add In Scope resources to validationResources HashTable'
    Write-Progress -Activity 'WARA Collector' -Status 'Adding In Scope Resources to Validation Resources' -PercentComplete 50 -Id 1
    foreach ($obj in $Scope_AllResources) {
        $key = "$($obj.id)"
        if (-not $validationResources.ContainsKey($key)) {
            $validationResources[$key] = $obj
        }
    }
    Write-Debug "Count of validationResourceIds: $($validationResources.count)"

    #Create validationResourceObj objects from the impactedResourceObj objects
    Write-Debug 'Creating validationResourceObj objects from the impactedResourceObj objects'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating Validation Resource Objects' -PercentComplete 53 -Id 1
    $validationResourceObj = Build-ValidationResourceObj -ValidationResources $validationResources -RecommendationObject $RecommendationObject -TypesNotInAPRLOrAdvisor $TypesNotInAPRLOrAdvisor
    Write-Debug "Count of validationResourceObj objects: $($validationResourceObj.count)"

    #Combine impactedResourceObj and validationResourceObj objects
    Write-Debug 'Combining impactedResourceObj and validationResourceObj objects'
    Write-Progress -Activity 'WARA Collector' -Status 'Combining Impacted and Validation Resource Objects' -PercentComplete 56 -Id 1
    $impactedResourceObj += $validationResourceObj
    Write-Debug "Count of combined validationResourceObj impactedResourceObj objects: $($impactedResourceObj.count)"

    #Get Advisor Metadata to include recommendations that are not in Advisor under 'HighAvailability'
    Write-Debug 'Getting Advisor Metadata'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Advisor Metadata' -PercentComplete 59 -Id 1
    try {
        $AdvisorMetadata = Get-WAFAdvisorMetadata -ResourceURL $BaseURL
    }
    catch {
        $AdvisorMetadata = @()
        $collectionErrors.Add([PSCustomObject]@{ Source = 'AdvisorMetadata'; Reachable = $false; Error = $_.Exception.Message })
        Write-Warning "Advisor metadata could not be retrieved in this environment; continuing without it. ($($_.Exception.Message))"
    }
    Write-Debug "Count of Advisor Metadata: $($AdvisorMetadata.count)"

    #Get Other Recommendations
    Write-Debug 'Getting Other Recommendations'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Other Recommendations' -PercentComplete 62 -Id 1
    $OtherRecommendations = if ($AdvisorMetadata) {
        Get-WARAOtherRecommendations -RecommendationObject $RecommendationObject -AdvisorMetadata $AdvisorMetadata
    }
    else { @() }
    Write-Debug "Count of Other Recommendations: $($OtherRecommendations.count)"

    #Get Advisor Recommendations
    Write-Debug 'Getting Advisor Recommendations'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Advisor Recommendations' -PercentComplete 65 -Id 1
    try {
        $advisorResourceObj = Get-WAFAdvisorRecommendation -AdditionalRecommendationIds $OtherRecommendations -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '') -HighAvailability
    }
    catch {
        $advisorResourceObj = @()
        $collectionErrors.Add([PSCustomObject]@{ Source = 'AdvisorRecommendations'; Reachable = $false; Error = $_.Exception.Message })
        Write-Warning "Advisor recommendations could not be retrieved in this environment; continuing without them. ($($_.Exception.Message))"
    }
    Write-Debug "Count of Advisor Recommendations: $($advisorResourceObj.count)"

    #Prior to filtering, capture all "global" recommendations that are microsoft.subscriptions/subscriptions since these get filtered out.
    Write-Debug 'Capturing global recommendations that are microsoft.subscriptions/subscriptions'
    Write-Progress -Activity 'WARA Collector' -Status 'Capturing Global Recommendations' -PercentComplete 68 -Id 1
    $globalRecommendations = $advisorResourceObj | Where-Object { $_.type -eq 'microsoft.subscriptions/subscriptions' }
    Write-Debug "Count of global recommendations: $($globalRecommendations.count)"

    #Filter Advisor Recommendations by subscription, resource group, and resource scope
    Write-Debug 'Filtering Advisor Recommendations by subscription, resource group, and resource scope'
    Write-Progress -Activity 'WARA Collector' -Status 'Filtering Advisor Recommendations' -PercentComplete 71 -Id 1
    $advisorResourceObj = Get-WAFFilteredResourceList -UnfilteredResources $advisorResourceObj.where({ $_.type -ne 'microsoft.subscriptions/subscriptions' }) -SubscriptionFilters $Scope_SubscriptionIds -ResourceGroupFilters $Scope_ResourceGroups
    Write-Debug "Count of filtered Advisor Recommendations: $($advisorResourceObj.count)"

    #If we passed tags, filter impactedResourceObj and advisorResourceObj by tagged resource group and tagged resource scope
    if (![string]::IsNullOrEmpty($Scope_Tags)) {
        Write-Debug 'Starting Tag Filtering'
        Write-Progress -Activity 'WARA Collector' -Status 'Starting Tag Filtering' -PercentComplete 72 -Id 1
        Write-Debug "Scope Tags: $Scope_Tags"

        #Get all tagged resource groups from the Implicit Subscription ID scope
        Write-Debug 'Getting all tagged resource groups from the Implicit Subscription ID scope'
        Write-Progress -Activity 'WARA Collector' -Status 'Getting Tagged Resource Groups' -PercentComplete 72 -Id 1
        $Filter_TaggedResourceGroupIds = Get-WAFTaggedResourceGroup -TagArray $Scope_Tags -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
        Write-Debug "Count of Tagged Resource Group Ids: $($Filter_TaggedResourceGroupIds.count)"

        #Get all tagged resources from the Implicit Subscription ID scope
        Write-Debug 'Getting all tagged resources from the Implicit Subscription ID scope'
        Write-Progress -Activity 'WARA Collector' -Status 'Getting Tagged Resources' -PercentComplete 73 -Id 1
        $Filter_TaggedResourceIds = Get-WAFTaggedResource -TagArray $Scope_Tags -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
        Write-Debug "Count of Tagged Resource Ids: $($Filter_TaggedResourceIds.count)"

        #Filter ResourceInventory objects by tagged resource group and resource scope
        Write-Debug 'Filtering ResourceInventory objects by tagged resource group and resource scope'
        Write-Progress -Activity 'WARA Collector' -Status 'Filtering Resource Inventory' -PercentComplete 73 -Id 1
        $ResourceInventory = Get-WAFFilteredResourceList -UnfilteredResources $ResourceInventory -ResourceGroupFilters $Filter_TaggedResourceGroupIds -ResourceFilters $Filter_TaggedResourceIds
        Write-Debug "Count of tag filtered ResourceInventory objects: $($ResourceInventory.count)"

        #Filter impactedResourceObj objects by tagged resource group and resource scope
        Write-Debug 'Filtering impactedResourceObj objects by tagged resource group and resource scope'
        Write-Progress -Activity 'WARA Collector' -Status 'Filtering Impacted Resource Objects' -PercentComplete 73 -Id 1
        $impactedResourceObj = Get-WAFFilteredResourceList -UnfilteredResources $impactedResourceObj -ResourceGroupFilters $Filter_TaggedResourceGroupIds -ResourceFilters $Filter_TaggedResourceIds
        Write-Debug "Count of tag filtered impactedResourceObj objects: $($impactedResourceObj.count)"

        #Filter Advisor Recommendations by tagged resource group and resource scope
        Write-Debug 'Filtering Advisor Recommendations by tagged resource group and resource scope'
        Write-Progress -Activity 'WARA Collector' -Status 'Filtering Advisor Recommendations' -PercentComplete 73 -Id 1
        $advisorResourceObj = Get-WAFFilteredResourceList -UnfilteredResources $advisorResourceObj -ResourceGroupFilters $Filter_TaggedResourceGroupIds -ResourceFilters $Filter_TaggedResourceIds
        Write-Debug "Count of tag filtered Advisor Recommendations: $($advisorResourceObj.count)"
    }

    #Build Specialized Resource Object if Specialized Workloads are selected but not present in the impactedResourceObj.
    #Some of the specialized workloads have queries that run. If this is the case then we need to check if the impactedResourceObj contains these resource types and if not add them to the impactedResourceObj.
    if ($SpecializedWorkloads) {
        Write-Debug 'Building Specialized Resource Object'

        Write-Progress -Activity 'WARA Collector' -Status 'Building Specialized Resource Object' -PercentComplete 74 -Id 1
        $specializedResourceObj = Build-SpecializedResourceObj -SpecializedResourceObj $SpecializedWorkloads -RecommendationObject $RecommendationObject
        Write-Debug "Count of Specialized Resource Object: $($specializedResourceObj.count)"

        Write-Debug 'Adding Specialized Resource Object to impactedResourceObj'
        $impactedResourceObj += $specializedResourceObj
        Write-Debug "Count of impactedResourceObj with Specialized Resource Object: $($impactedResourceObj.count)"
    }

    #Add global recommendations back to advisorResourceObj
    Write-Debug 'Adding global recommendations back to advisorResourceObj'
    Write-Progress -Activity 'WARA Collector' -Status 'Adding Global Recommendations' -PercentComplete 75 -Id 1
    $advisorResourceObj += $globalRecommendations
    Write-Debug "Count of advisorResourceObj with global recommendations: $($advisorResourceObj.count)"

    #Build Resource Type Object
    Write-Debug 'Building Resource Type Object with impactedResourceObj and advisorResourceObj'
    Write-Progress -Activity 'WARA Collector' -Status 'Building Resource Type Object' -PercentComplete 78 -Id 1
    $resourceTypeObj = Build-ResourceTypeObj -ResourceObj $Scope_AllResources -TypesNotInAPRLOrAdvisor $TypesNotInAPRLOrAdvisor
    Write-Debug "Count of Resource Type Object : $($resourceTypeObj.count)"

    #Get Azure Outages
    Write-Debug 'Getting Azure Outages'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Azure Outages' -PercentComplete 81 -Id 1
    try {
        $outageResourceObj = Get-WAFOldOutage -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
    }
    catch {
        $outageResourceObj = @()
        $collectionErrors.Add([PSCustomObject]@{ Source = 'Outages'; Reachable = $false; Error = $_.Exception.Message })
        Write-Warning "Azure outages (Microsoft.ResourceHealth) could not be retrieved; continuing without them. ($($_.Exception.Message))"
    }

    #Get Azure Retirements
    Write-Debug 'Getting Azure Retirements'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Azure Retirements' -PercentComplete 84 -Id 1
    try {
        $retirementResourceObj = Get-WAFResourceRetirement -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
    }
    catch {
        $retirementResourceObj = @()
        $collectionErrors.Add([PSCustomObject]@{ Source = 'Retirements'; Reachable = $false; Error = $_.Exception.Message })
        Write-Warning "Azure retirements (Microsoft.ResourceHealth) could not be retrieved; continuing without them. ($($_.Exception.Message))"
    }

    #Get Azure Support Tickets
    Write-Debug 'Getting Azure Support Tickets'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Azure Support Tickets' -PercentComplete 87 -Id 1
    try {
        $supportTicketObjects = Get-WAFSupportTicket -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
    }
    catch {
        $supportTicketObjects = @()
        $collectionErrors.Add([PSCustomObject]@{ Source = 'SupportTickets'; Reachable = $false; Error = $_.Exception.Message })
        Write-Warning "Azure support tickets (Microsoft.Support) could not be retrieved; continuing without them. ($($_.Exception.Message))"
    }

    #Get Azure Service Health
    Write-Debug 'Getting Azure Service Health'
    Write-Progress -Activity 'WARA Collector' -Status 'Getting Azure Service Health' -PercentComplete 90 -Id 1
    try {
        $serviceHealthObjects = Get-WAFServiceHealth -SubscriptionIds $Scope_ImplicitSubscriptionIds.replace('/subscriptions/', '')
    }
    catch {
        $serviceHealthObjects = @()
        $collectionErrors.Add([PSCustomObject]@{ Source = 'ServiceHealth'; Reachable = $false; Error = $_.Exception.Message })
        Write-Warning "Azure service health alerts could not be retrieved; continuing without them. ($($_.Exception.Message))"
    }

    $stopWatch.Stop()
    Write-Debug "Elapsed Time: $($stopWatch.Elapsed.toString('hh\:mm\:ss'))"

    #Create Script Details Object
    Write-Debug 'Creating Script Details Object'
    $scriptDetails = [PSCustomObject]@{
        Version                        = $(Get-Module -Name $MyInvocation.MyCommand.ModuleName).Version.toString()
        ElapsedTime                    = $stopWatch.Elapsed.toString('hh\:mm\:ss')
        SAP                            = [bool]$SAP
        AVD                            = [bool]$AVD
        AVS                            = [bool]$AVS
        HPC                            = [bool]$HPC
        AI_GPT_RAG                     = [bool]$AI_GPT_RAG
        ORACLE                         = [bool]$ORACLE
        TenantId                       = $Scope_TenantId
        SubscriptionIds                = $Scope_SubscriptionIds
        ResourceGroups                 = $Scope_ResourceGroups
        ImplicitSubscriptionIds        = $Scope_ImplicitSubscriptionIds
        Tags                           = $Scope_Tags
        AzureEnvironment               = $AzureEnvironment
        RecommendationDataUri          = $RecommendationDataUri
        RecommendationResourceTypesUri = $RecommendationResourceTypesUri
        UseImplicitRunbookSelectors    = $UseImplicitRunbookSelectors
        RunbookFile                    = $RunbookFile
        ConfigFile                     = $ConfigFile
        ConfigData                     = $ConfigData
        RunTimeParameters              = $scriptParams
    }

    #Create output JSON
    Write-Debug 'Creating output JSON'
    Write-Progress -Activity 'WARA Collector' -Status 'Creating Output JSON' -PercentComplete 93 -Id 1
    $outputJson = [PSCustomObject]@{
        scriptDetails       = $scriptDetails
        impactedResources   = $impactedResourceObj
        resourceType        = $resourceTypeObj
        advisory            = $advisorResourceObj
        outages             = $outageResourceObj
        retirements         = $retirementResourceObj
        supportTickets      = $supportTicketObjects
        serviceHealth       = $serviceHealthObjects
        resourceInventory   = $ResourceInventory
        collectionErrors    = $collectionErrors
    }

    Write-Debug 'Output JSON'
    Write-Progress -Activity 'WARA Collector' -Status 'Output Data' -PercentComplete 100 -Id 1 -Completed

    #Return data if PassThru is enabled
    if ($PassThru) {
        Write-host "Returning data as PassThru is enabled." -ForegroundColor Yellow
        return $outputJson
    }

    #Output JSON to file
    $outputPath = ('.\WARA-File-' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '.json')
    Write-Host "Output Path: $outputPath" -ForegroundColor Yellow
    $outputJson | ConvertTo-Json -Depth 15 | Out-file $outputPath
}

function Build-ImpactedResourceObj {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject] $ImpactedResources,

        [Parameter(Mandatory = $true)]
        [Hashtable] $AllResources,

        [Parameter(Mandatory = $true)]
        [Hashtable] $RecommendationObject
    )

    $impactedResourceObj = [impactedResourceFactory]::new($ImpactedResources, $AllResources, $RecommendationObject)
    $r = $impactedResourceObj.createImpactedResourceObjects()

    return , $r
}

function Build-ValidationResourceObj {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Hashtable] $ValidationResources,

        [Parameter(Mandatory = $true)]
        [PSObject] $RecommendationObject,

        [Parameter(Mandatory = $true)]
        [PSObject] $TypesNotInAPRLOrAdvisor
    )

    $validatorObj = [validationResourceFactory]::new($RecommendationObject, $validationResources, $TypesNotInAPRLOrAdvisor)
    $r = $validatorObj.createValidationResourceObjects()

    return , $r
}

function Build-ResourceTypeObj {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSObject] $ResourceObj,

        [Parameter(Mandatory = $true)]
        [PSObject] $TypesNotInAPRLOrAdvisor
    )

    $return = [resourceTypeFactory]::new($ResourceObj, $TypesNotInAPRLOrAdvisor).createResourceTypeObjects()

    return , $return
}

function Build-SpecializedResourceObj {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSObject] $SpecializedResourceObj,

        [Parameter(Mandatory = $true)]
        [PSObject] $RecommendationObject
    )

    $return = [specializedResourceFactory]::new($SpecializedResourceObj, $RecommendationObject).createSpecializedResourceObjects()

    return , $return
}

function Get-WARAOtherRecommendations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject] $RecommendationObject,

        [Parameter(Mandatory = $true)]
        [PSObject] $AdvisorMetadata
    )

    $metadata = $AdvisorMetadata.where({ $_.recommendationCategory -ne 'HighAvailability' }).id

    #Returns recommendations that are in APRL but not in Advisor under 'HighAvailability'
    $return = $RecommendationObject.recommendationTypeId | Where-Object { $_ -in $metadata }

    return , $return
}


<#
.CLASS
    aprlResourceTypeObj

.SYNOPSIS
    Represents a resource type object for APRL.

.DESCRIPTION
    The `aprlResourceTypeObj` class encapsulates the details of a resource type in APRL, including the number of resources, availability in APRL/ADVISOR, assessment owner, status, and notes.

.PROPERTY Resource Type
    The type of the resource.

.PROPERTY Number Of Resources
    The number of resources of this type.

.PROPERTY Available in APRL/ADVISOR?
    Indicates whether the resource type is available in APRL or ADVISOR.

.PROPERTY Assessment Owner
    The owner of the assessment.

.PROPERTY Status
    The status of the resource type.

.PROPERTY Notes
    Additional notes about the resource type.

.EXAMPLE
    $resourceType = [aprlResourceTypeObj]::new()
    $resourceType.'Resource Type' = "Microsoft.Compute/virtualMachines"
    $resourceType.'Number Of Resources' = 5

.NOTES
    Author: Kyle Poineal
    Date: 2023-10-07
#>
class aprlResourceTypeObj {
    [string] ${Resource Type}
    [int] ${Number Of Resources}
    [string] ${Available in APRL/ADVISOR?}
    [string] ${Assessment Owner}
    [string] $Status
    [string] $Notes
}

<#
.CLASS
    resourceTypeFactory

.SYNOPSIS
    Factory class to create resource type objects.

.DESCRIPTION
    The `resourceTypeFactory` class is responsible for creating instances of `aprlResourceTypeObj` based on impacted resources and types not in APRL or ADVISOR.

.PROPERTY impactedResourceObj
    The impacted resource objects grouped by type.

.PROPERTY TypesNotInAPRLOrAdvisor
    Resource types that are not available in APRL or ADVISOR.

.CONSTRUCTORS
    resourceTypeFactory([PSObject]$impactedResourceObj, [PSObject]$TypesNotInAPRLOrAdvisor)
        Initializes a new instance of the `resourceTypeFactory` class.

.METHODS
    [object[]] createResourceTypeObjects()
        Creates and returns an array of `aprlResourceTypeObj` instances based on the impacted resources.

.EXAMPLE
    $factory = [resourceTypeFactory]::new($impactedResourceObj, $TypesNotInAPRLOrAdvisor)
    $resourceTypes = $factory.createResourceTypeObjects()

.NOTES
    Author: Kyle Poineal
    Date: 2023-10-07
#>
class resourceTypeFactory {
    [PSObject]$impactedResourceObj
    [PSObject]$TypesNotInAPRLOrAdvisor

    resourceTypeFactory([PSObject]$impactedResourceObj, [PSObject]$TypesNotInAPRLOrAdvisor) {
        $this.impactedResourceObj = $impactedResourceObj | Group-Object -Property type | Select-Object Name, Count
        $this.TypesNotInAPRLOrAdvisor = $TypesNotInAPRLOrAdvisor
    }

    [object[]] createResourceTypeObjects() {
        $return = foreach ($type in $this.impactedResourceObj) {
            $r = [aprlResourceTypeObj]::new()
            $r.'Resource Type' = $type.Name
            $r.'Number Of Resources' = $type.Count
            $r.'Available in APRL/ADVISOR?' = $(($this.TypesNotInAPRLOrAdvisor -contains $type.Name) ? "No" : "Yes")
            $r.'Assessment Owner' = ""
            $r.Status = ""
            $r.notes = ""

            $r
        }
        return $return
    }
}

<#
.CLASS
    aprlResourceObj

.SYNOPSIS
    Represents an APRL resource object.

.DESCRIPTION
    The `aprlResourceObj` class encapsulates the details of an APRL resource, including validation action, recommendation ID, name, ID, type, location, subscription ID, resource group, parameters, check name, and selector.

.PROPERTY validationAction
    The validation action for the resource.

.PROPERTY recommendationId
    The recommendation ID for the resource.

.PROPERTY name
    The name of the resource.

.PROPERTY id
    The ID of the resource.

.PROPERTY type
    The type of the resource.

.PROPERTY location
    The location of the resource.

.PROPERTY subscriptionId
    The subscription ID of the resource.

.PROPERTY resourceGroup
    The resource group of the resource.

.PROPERTY param1
    Additional parameter 1.

.PROPERTY param2
    Additional parameter 2.

.PROPERTY param3
    Additional parameter 3.

.PROPERTY param4
    Additional parameter 4.

.PROPERTY param5
    Additional parameter 5.

.PROPERTY checkName
    The check name for the resource.

.PROPERTY selector
    The selector for the resource.

.EXAMPLE
    $resource = [aprlResourceObj]::new()
    $resource.name = "myResource"
    $resource.id = "/subscriptions/12345/resourceGroups/myRG/providers/Microsoft.Compute/virtualMachines/myVM"

.NOTES
    Author: Kyle Poineal
    Date: 2023-10-07
#>
class aprlResourceObj {
    [string] $validationAction
    [string] $recommendationId
    [string] $name
    [string] $id
    [string] $type
    [string] $location
    [string] $subscriptionId
    [string] $resourceGroup
    [string] $param1
    [string] $param2
    [string] $param3
    [string] $param4
    [string] $param5
    [string] $checkName
    [string] $selector
}

<#
.CLASS
    impactedResourceFactory

.SYNOPSIS
    Factory class to create impacted resource objects.

.DESCRIPTION
    The `impactedResourceFactory` class is responsible for creating instances of `aprlResourceObj` based on impacted resources, all resources, and recommendation objects.

.PROPERTY impactedResources
    The impacted resources.

.PROPERTY allResources
    All resources.

.PROPERTY RecommendationObject
    The recommendation object.

.CONSTRUCTORS
    impactedResourceFactory([PSObject]$impactedResources, [hashtable]$allResources, [hashtable]$RecommendationObject)
        Initializes a new instance of the `impactedResourceFactory` class.

.METHODS
    [object[]] createImpactedResourceObjects()
        Creates and returns an array of `aprlResourceObj` instances based on the impacted resources.

.EXAMPLE
    $factory = [impactedResourceFactory]::new($impactedResources, $allResources, $RecommendationObject)
    $impactedResources = $factory.createImpactedResourceObjects()

.NOTES
    Author: Kyle Poineal
    Date: 2023-10-07
#>
class impactedResourceFactory {
    [PSObject] $impactedResources
    [hashtable] $allResources
    [hashtable] $RecommendationObject

    impactedResourceFactory([PSObject]$impactedResources, [hashtable]$allResources, [hashtable]$RecommendationObject) {
        $this.impactedResources = $impactedResources
        $this.allResources = $allResources
        $this.RecommendationObject = $RecommendationObject
    }

    [object[]] createImpactedResourceObjects() {
        $return = foreach ($impactedResource in $this.impactedResources) {
            $r = [aprlResourceObj]::new()
            $r.validationAction = "APRL - Queries"
            $r.RecommendationId = $impactedResource.recommendationId
            $r.Name = $impactedResource.name
            $r.Id = $impactedResource.id
            $r.type = $this.RecommendationObject[$r.recommendationId].recommendationResourceType ?? $this.allResources[$r.id].type ?? "Unknown"
            $r.location = $this.allResources[$r.id].location ?? "Unknown"
            $r.subscriptionId = $this.allResources[$r.id].subscriptionId ?? $r.id.split("/")[2] ?? "Unknown"
            $r.resourceGroup = $this.allResources[$r.id].resourceGroup ?? $r.id.split("/")[4] ?? "Unknown"
            $r.Param1 = $impactedResource.param1
            $r.Param2 = $impactedResource.param2
            $r.Param3 = $impactedResource.param3
            $r.Param4 = $impactedResource.param4
            $r.Param5 = $impactedResource.param5
            $r.checkName = $impactedResource.checkName
            $r.selector = $impactedResource.selector ?? "APRL"
            $r
        }
        return $return
    }
}

<#
.CLASS
    validationResourceFactory

.SYNOPSIS
    Factory class to create validation resource objects.

.DESCRIPTION
    The `validationResourceFactory` class is responsible for creating instances of `aprlResourceObj` for validation purposes based on recommendation objects, validation resources, and types not in APRL or ADVISOR.

.PROPERTY recommendationObject
    The recommendation object.

.PROPERTY validationResources
    The validation resources.

.PROPERTY TypesNotInAPRLOrAdvisor
    Resource types that we want to create a recommendation for but do not have a recommendation for.

.CONSTRUCTORS
    validationResourceFactory([PSObject]$recommendationObject, [hashtable]$validationResources, [PSObject]$TypesNotInAPRLOrAdvisor)
        Initializes a new instance of the `validationResourceFactory` class.

.METHODS
    [object[]] createValidationResourceObjects()
        Creates and returns an array of `aprlResourceObj` instances for validation purposes.

    static [string] getValidationAction($query)
        Determines the validation action based on the query.

.EXAMPLE
    $factory = [validationResourceFactory]::new($recommendationObject, $validationResources, $TypesNotInAPRLOrAdvisor)
    $validationResources = $factory.createValidationResourceObjects()

.NOTES
    Author: Kyle Poineal
    Date: 2023-10-07
#>
class validationResourceFactory {
    # This class is used to create validationResourceObj objects

    # Properties
    [PSObject] $recommendationObject # The recommendation object
    [hashtable] $validationResources # The validation resources
    [PSObject] $TypesNotInAPRLOrAdvisor # Resource types that we want to create a recommendation for but do not have a recommendation for.

    validationResourceFactory([PSObject]$recommendationObject, [hashtable]$validationResources, [PSObject]$TypesNotInAPRLOrAdvisor) {
        $this.recommendationObject = $recommendationObject
        $this.validationResources = $validationResources
        $this.TypesNotInAPRLOrAdvisor = $TypesNotInAPRLOrAdvisor
    }

    [object[]] createValidationResourceObjects() {
        $return = @()

        $RecommendationResourceTypesWithoutAutomation = ($this.recommendationObject | group recommendationResourceType | where {$_.group.automationavailable -notcontains $false}).Name

        $return = foreach ($v in $this.validationResources.GetEnumerator()) {

            $impactedResource = $v.value

            if($impactedResource.type -in $RecommendationResourceTypesWithoutAutomation){
                continue
            }

            $recommendationByType = $this.recommendationObject.where({ $_.automationAvailable -eq $false -and $impactedResource.type -eq $_.recommendationResourceType -and $_.recommendationMetadataState -eq "Active" -and [string]::IsNullOrEmpty($_.recommendationTypeId) })

            if ($recommendationByType) {
                foreach ($rec in $recommendationByType) {
                    $r = [aprlResourceObj]::new()
                    $r.validationAction = [validationResourceFactory]::getValidationAction($rec.query)
                    $r.recommendationId = $rec.aprlGuid
                    $r.name = $impactedResource.name
                    $r.id = $impactedResource.id
                    $r.type = $impactedResource.type
                    $r.location = $impactedResource.location
                    $r.subscriptionId = $impactedResource.subscriptionId
                    $r.resourceGroup = $impactedResource.resourceGroup
                    $r.param1 = ''
                    $r.param2 = ''
                    $r.param3 = ''
                    $r.param4 = ''
                    $r.param5 = ''
                    $r.checkName = ''
                    $r.selector = $impactedResource.selector ?? "APRL"
                    $r
                }
            }
            elseif ($impactedResource.type -in $this.TypesNotInAPRLOrAdvisor) {
                $r = [aprlResourceObj]::new()
                $r.validationAction = [validationResourceFactory]::getValidationAction("No Recommendations")
                $r.recommendationId = ''
                $r.name = $impactedResource.name
                $r.id = $impactedResource.id
                $r.type = $impactedResource.type
                $r.location = $impactedResource.location
                $r.subscriptionId = $impactedResource.subscriptionId
                $r.resourceGroup = $impactedResource.resourceGroup
                $r.param1 = ''
                $r.param2 = ''
                $r.param3 = ''
                $r.param4 = ''
                $r.param5 = ''
                $r.checkName = ''
                $r.selector = $impactedResource.selector ?? "APRL"
                $r
            }
            else {
                Write-Error "No recommendation found for $($impactedResource.type) with resource id $($impactedResource.id)"
            }
        }

        return $return
    }

    static [string] getValidationAction($query) {
        $return = switch -wildcard ($query) {
            "*development*" { 'IMPORTANT - Query under development - Validate Resources manually' }
            "*cannot-be-validated-with-arg*" { 'IMPORTANT - Recommendation cannot be validated with ARGs - Validate Resources manually' }
            "*Azure Resource Graph*" { 'IMPORTANT - Query under development - Validate Resources manually' }
            "No Recommendations" { 'IMPORTANT - Resource Type is not available in either APRL or Advisor - Validate Resources manually if applicable, if not delete this line' }
            default { 'IMPORTANT - Recommendation cannot be validated with ARGs - Validate Resources manually' }
            #default { "IMPORTANT - Query does not exist - Validate Resources Manually" }
        }
        return $return
    }
}

<#
.CLASS
    specializedResourceFactory

.SYNOPSIS
    Factory class to create specialized resource objects.

.DESCRIPTION
    The `specializedResourceFactory` class is responsible for creating instances of `aprlResourceObj` for specialized resources based on recommendation objects.

.PROPERTY specializedResources
    The specialized resources.

.PROPERTY RecommendationObject
    The recommendation object.

.CONSTRUCTORS
    specializedResourceFactory([PSObject]$specializedResources, [PSObject]$RecommendationObject)
        Initializes a new instance of the `specializedResourceFactory` class.

.METHODS
    [object[]] createSpecializedResourceObjects()
        Creates and returns an array of `aprlResourceObj` instances for specialized resources.

    static [string] getValidationAction($query)
        Determines the validation action based on the query.

.EXAMPLE
    $factory = [specializedResourceFactory]::new($specializedResources, $RecommendationObject)
    $specializedResources = $factory.createSpecializedResourceObjects()

.NOTES
    Author: Kyle Poineal
    Date: 2023-10-07
#>
class specializedResourceFactory {
    # This class is used to create specializedResourceObj objects

    # Properties
    [PSObject] $specializedResources # The specialized resources
    [PSObject] $RecommendationObject # The recommendation object

    specializedResourceFactory([PSObject]$specializedResources, [PSObject]$RecommendationObject) {
        $this.specializedResources = $specializedResources
        $this.RecommendationObject = $RecommendationObject
    }

    [object[]] createSpecializedResourceObjects() {
        $return = foreach ($s in $this.specializedResources) {

            $thisType = $this.RecommendationObject.where({ $s -in $_.tags -and $_.recommendationMetadataState -eq "Active" })
            foreach ($type in $thisType) {
                $r = [aprlResourceObj]::new()
                $r.validationAction = [specializedResourceFactory]::getValidationAction($type.query)
                $r.recommendationId = $type.aprlGuid
                $r.name = ''
                $r.id = ''
                $r.type = $type.recommendationResourceType
                $r.location = ''
                $r.subscriptionId = ''
                $r.resourceGroup = ''
                $r.param1 = ''
                $r.param2 = ''
                $r.param3 = ''
                $r.param4 = ''
                $r.param5 = ''
                $r.checkName = ''
                $r.selector = "APRL"
                $r
            }
        }
        return $return
    }

    static [string] getValidationAction($query) {
        $return = switch -wildcard ($query) {
            "*development*" { 'IMPORTANT - Query under development - Validate Resources manually' }
            "*cannot-be-validated-with-arg*" { 'IMPORTANT - Recommendation cannot be validated with ARGs - Validate Resources manually' }
            "*Azure Resource Graph*" { 'IMPORTANT - Query under development - Validate Resources manually' }
            "No Recommendations" { 'IMPORTANT - Resource Type is not available in either APRL or Advisor - Validate Resources manually if applicable, if not delete this line' }
            default { 'IMPORTANT - Recommendation cannot be validated with ARGs - Validate Resources manually' }
            #default { "IMPORTANT - Query does not exist - Validate Resources Manually" }
        }
        return $return
    }
}

# SIG # Begin signature block
# MIIoLwYJKoZIhvcNAQcCoIIoIDCCKBwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAWnspEvyToWeXa
# VPJmIo/4NLYYQUmfz3zmz+Ss1JJSzKCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIEOKmzHL/PgyAJVG/vfjOYBl
# MmN6ll3bsJl7v7XmeY/IMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCBW0qsy8/lDYjxpfFh8hD4npkLh+jP4jyq6zcXI7xaUCQHlE63Wuxi
# CCF25fhJSn1+EW9vaQQCy9CtBeTdAaNAoBbHTn+S3PnNUn6vpnwAuXzceQjpLLEe
# hGMpbw12jJaUpp91hi5NyJCYpY/f/ONbsSl6fSNaqKwlLGdzs2Ut2j1juS96SZtU
# yB1PyQ47Y/25SU1ydPNfmPMobJVPcl1EkTa6ON0IBm68Wx9i4S1Gc6Nbh8zhDIMA
# fibmDgLiSIGrL2kLmcCz7lhbR/V9ebQ7wADtnFgDZf1cscltCHaZf4XrFw0vZJbZ
# HxFnKRYE5zdp4XDN1KFZw0mxHJ7W2tCKoYIXlzCCF5MGCisGAQQBgjcDAwExgheD
# MIIXfwYJKoZIhvcNAQcCoIIXcDCCF2wCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIPYkEuLD0zv7wyPY4SjtqBCFTDpQf97WSgN/cQReiL2IAgZoJmAf
# wuIYEzIwMjUwNTIyMTUwMDAxLjcwNFowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpGMDAy
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEe0wggcgMIIFCKADAgECAhMzAAACBTx1bIJEh83+AAEAAAIFMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI0OVoXDTI2MDQyMjE5NDI0OVowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpGMDAyLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAJKS8t93uAXWvbAZ3LzkIVhcdQLzLATIPgtEu/RZ
# gRd55nS7li4runZxdXNCZk84dM3xNaobZI+VRv7s+V3MUMMCVHe9TymI6OaG72sd
# bjczZ1uiv2OX2CW6HPBR3ZJyJUZrt/23ru0zcoUpFIxcW0coXcrAEHtpWj5vWrLm
# n0NaWjY3kUasGocwRWU3LOt4digMsv6bx9Kyoy7+JhSrHMrkXhLshjk16YAHwH4D
# dCDBUiLrgokh94plR6JYoJN/ih1SA/cBCKXGHjw/rPsBPggDR0wS0qFgsWzhb8o0
# MyxivsqlA8pUJwLkTy4Md0p7C5vLN6eRPHLh9/U3eDzKGjk+L0F+NHRXK2uSakN9
# 6nwk/BxvgE04hc6jWl90dnwS+dHskVkVCKqkxkWU7kIC5Ngfy6Nzk9QeVowAnw0R
# r1MUlM5IGsHs9GB6H/o0nbG73LE+H+RU30Eayz3cwLpSOmjF4zjHvRvBvCIrxI0c
# g7wPxyqXtVJ69RhuM3g2iAUXCEEKWGh0T/N4Y+rrLqLEPjPrkgdjfPAVBsFVf/D4
# v9Uc2f8EwazY8YeeVGM78qTw0ik3iyGQVCoDV8zTx+usNI0Rj1UoO2mxSkAXnVjW
# hDq0mFYzV3ed4JeHpv/o35d4c5ELCdjzcr6kUwyGDyKxdBvopGXrmSDxdgF3gnCR
# sqhVAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUdHpauzbtvr7IndsRn2jk10vVsvAw
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAA67GTjtvWLvlXzVPHrGjXTE0ivjNnFh
# V+QXlMWraSd08eDNXIueyzD8cqQRlEKBlWmQoQnpjiO8bm26AyL5aO3uFQxKKmT5
# GkHmAVC+HXGYAQvZ+V6DNNYSyePTCsRKmUjlne+B/Z4ZcCv9FyoaNKmi4dsPYdj1
# jcXQ6XoVEMJX8cQYpEfOfYzmtCkUZKNpxPgOSpViZ6b8Cs59K9WiOcoQhb3XhTEa
# 26ElKv2M6jlGpNfsYipu283vOFaV36LdfF2F01x+VDP1iGgWpB7EF6eghGAA8C3A
# frzFzOv0swLeX0AsSmey87RGIo9cXiXbb859wV99qmGeX9MPCSIl/E7IAx9QXfEj
# 37eLNPVZfYIWzZFo4Crd1yiHInD7FEbQTzCQIeqRXnsFtQETMtfwv2UYnUOjFg2m
# gNfuJDMn1B2TKmLl+/vvcTcKHwD62jF4WnoWJ9BnIJzwhwgGUfJqToxtPphNVA2B
# D+HwUX5Lk6o/sIQIW8gfYq3Y/QaU670LQF9qyGdOOh2a1kVmym1S+KuT/Yc4suMG
# IyMPAmxAAUDMm7Phm1PKiuu8RHXQaX+ZuZWIeqG80AJ9PM+It+MODK+n2zv9se78
# JqUqZ6IlQsaOH+XACPnfX2mCCwYmpuEngVsVz6hTnN99ub+sqNVnyTE0PeKbXRCn
# jWfa3+qnI/oRMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046RjAwMi0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# ANWwf2nW0mf6SvAIy+o6FW9ett60oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2W/fMCIYDzIwMjUwNTIyMDk0
# MTUxWhgPMjAyNTA1MjMwOTQxNTFaMHcwPQYKKwYBBAGEWQoEATEvMC0wCgIFAOvZ
# b98CAQAwCgIBAAICD3UCAf8wBwIBAAICE3wwCgIFAOvawV8CAQAwNgYKKwYBBAGE
# WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDAN
# BgkqhkiG9w0BAQsFAAOCAQEAr3LLfLckdZ0OxANPHBCyg98jDAYpiERuACVaoHAF
# 0zD9yI6VqmQnmgLUZN5WbKDPUvjByzUrT/4t0BtPzr9OFOE1Dhcr8VsuJ9ea6ODi
# s/V0JXT0/nBD/2Ha/WVt8jMRe2+gTZFWJW0U9ljDaOaoL534/oPUziixZ/R9EhhK
# adAZacsdKYdbiOg1lFGB4Cn2zoF/PBKEs4jlsrw8ognGXKNEoE8N40xEAQMQM5GQ
# NysX6y83G8FKHGbvPPEHOoXJG17KrOQR9AJiHKDUJ7xCgGk6nDX7sB8cBwLKgHwS
# RLAuT/HISx4PVkmQsvutwpy3ZW74DSPwneQS7QOhdanEBDGCBA0wggQJAgEBMIGT
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACBTx1bIJEh83+AAEA
# AAIFMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQ
# AQQwLwYJKoZIhvcNAQkEMSIEIBOSPjPXFc+WeVKhsSuKJp3Ei3BUmVcWHTrXHpcs
# xtWuMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQggA0DPdx5jS6aF1YtHawm
# mrQ4+q0kNMBhGaMdWTARb0gwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAgU8dWyCRIfN/gABAAACBTAiBCDMrlWpoNDvr9UHhLT6XWKy
# gfzQJBK1YnHX02R4ugGTPTANBgkqhkiG9w0BAQsFAASCAgAwuggvPbxQyLaHFw76
# SDm1/2sXaHZ+VEsvndpJuasKG1Q7cdu+OZyzfNqEQHaEDmfytTsS0Qdu8vwRkLay
# OfO5K1c9JO2X/qrPZekSXmwVgnnUhBbLM2SJedknHLFzUFCWQRyoSG+kxwfmxRgT
# vwUbiujWBMLY5nYXZTMh7nPN+zLr5ux5CE5OUOs4JfLKxNvgM7DiCk8PpmdhFLBj
# bYi2KcOBoJTcwpK6F4NI+A7WLdd0bF+h/HshdJoYxaFMvc/b3ysxONkZ2p+hgJ3g
# sch/Griit04FBBHQNl2dJZltHYQCKZqWHtPbGrTtlkvLAb9KjkWdTEs0xpiY7qic
# rR4wUXMTF7Tl32e8TqiUTWu7L3s53g3DDyJkY7j+HQD1A7c723/Xb5i4luvAI+m1
# ZklvFRro6RV4Z3pEWQ3KLAsZEVWD20mcmSyIzEt3Wya0uo+Q45cU76ejP8aVKV1S
# lGHJN0fdxL+B04L2isnTnVwWJhVJhSgCRuIk+/5yK1SRcjsGU1vu4fFAcCV6IRM2
# H94Ov46FRwbwv1t2Nxu3iZLXsKSwOsj+MvTQ49U7Es6EJ+KIlGcMRuYf9Da439ke
# Xm0U4glD6PjZBSCULYbSLD969qgWqZzEeEUhMA+JekguDukTE1DKEzW6W4ExIYsR
# 8B8BY2rXaSrkpmn11j0LBGLbYQ==
# SIG # End signature block
