#Requires -Version 7

<#
.SYNOPSIS
Well-Architected Reliability Assessment Report Generator Script

.DESCRIPTION
The script "3_wara_reports_generator" processes the Excel file created by the "2_wara_data_analyzer" script and generates the final PowerPoint and Word reports for the Well-Architected Reliability Assessment.

.PARAMETER Help
Switch to display help information.

.PARAMETER CustomerName
Name of the customer for whom the report is being generated.

.PARAMETER WorkloadName
Name of the workload being assessed.

.PARAMETER ExcelFile
Path to the Excel file created by the "2_wara_data_analyzer" script.

.PARAMETER Heavy
Switch to enable heavy processing mode. When enabled, this mode introduces additional delays using Start-Sleep at various points in the script to handle heavy environments more gracefully. This can help in scenarios where the system resources are limited or the operations being performed are resource-intensive, ensuring the script doesn't overwhelm the system.

.PARAMETER PPTTemplateFile
Path to the PowerPoint template file.

.EXAMPLE
.\3_wara_reports_generator.ps1 -ExcelFile 'C:\WARA_Script\WARA Action Plan 2024-03-07_16_06.xlsx' -CustomerName 'ABC Customer' -WorkloadName 'SAP On Azure' -Heavy -PPTTemplateFile 'C:\Templates\Template.pptx' -WordTemplateFile 'C:\Templates\Template.docx'

.LINK
https://github.com/Azure/Azure-Proactive-Resiliency-Library-v2
#>

Param(
  [switch] $Help,
  #[switch] $csvExport,
  [switch] $includeLow,
  [string] $CustomerName,
  [string] $WorkloadName,
  [Parameter(mandatory = $true)]
  [Alias('ExcelFile')]
  [string] $ExpertAnalysisFile,
  [string] $AssessmentFindingsFile,
  [string] $PPTTemplateFile
)

# Checking the operating system running this script.
if (-not $IsWindows) {
  Write-Host 'This script only supports Windows operating systems currently. Please try to run with Windows operating systems.'
  Exit
  }

  # Check if Clipboard History is enabled
$clipboardHistory = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -ErrorAction SilentlyContinue

if ($clipboardHistory.EnableClipboardHistory -eq 1) {
    Throw "Clipboard History is enabled. Please disable Clipboard History before running this script."
} else {
    Write-Debug "Clipboard History is disabled."
}

# TODO: Remove if not needed
<# $CurrentPath = Get-Location
$CurrentPath = $CurrentPath.Path #>
if ($PSBoundParameters.ContainsKey('PPTTemplateFile')) {
    # Resolve-Path throw exception if the path does not exist.
    $pptTemplateFilePath = (Resolve-Path -LiteralPath $PPTTemplateFile).Path
    if (-not (Test-Path -PathType Leaf -LiteralPath $pptTemplateFilePath)) {
        # The specified path is not a file, it may be a folder.
        Write-Error -Message ('The specified PowerPoint template file "{0}" is not a file. Please provide the path to the PowerPoint template file.' -f $pptTemplateFilePath)
        Exit  # TODO: This can be deleted after adding exception handling.
    }
}
else {
    $pptTemplateFilePath = Join-Path -Path $PSScriptRoot -ChildPath 'Mandatory - Executive Summary presentation - Template.pptx'
    if (-not (Test-Path -PathType Leaf -LiteralPath $pptTemplateFilePath)) {
        Write-Error -Message ('The default PowerPoint template file "{0}" does not exist. Please contact the WARA team via GitHub or Microsoft Teams.' -f $pptTemplateFilePath)  # TODO
        Exit  # TODO: This can be deleted after adding exception handling.
    }
}
Write-Host ('PowerPoint Template File: {0}' -f $pptTemplateFilePath)

if (!$AssessmentFindingsFile) {
  write-host ("$PSScriptRoot/Assessment-Findings-Report-v1.xlsx")
  if ((Test-Path -Path ("$PSScriptRoot/Assessment-Findings-Report-v1.xlsx") -PathType Leaf) -eq $true) {
    $AssessmentFindingsFile = ("$PSScriptRoot/Assessment-Findings-Report-v1.xlsx")
  }
  else {
    Write-Error "Assessment Findings file is missing. Please provide the path to the Assessment Findings file."
    Exit
  }
}

if (!$ExpertAnalysisFile) {
  Write-Host "The Expert-Analysis Excel file is missing. Please provide the path to the Expert-Analysis Excel file." -ForegroundColor Yellow
  Exit
}


if (!$CustomerName) {
  $CustomerName = '[Customer Name]'
}

if (!$WorkloadName) {
  $WorkloadName = '[Workload Name]'
}

$TableStyle = 'Light19'

# Excel COM helpers (Microsoft-only) that replace the ImportExcel/EPPlus operations for Phase 3.
. "$PSScriptRoot/WARAReportExcelCom.ps1"

  ######################## REGULAR Functions ##########################

  function Test-ReviewedRecommendations {
    Param($ExcelFile)

    $ExcelContent = Import-WARAExcelViaCom -Path $ExcelFile -WorksheetName '4.ImpactedResourcesAnalysis' -StartRow 12

    if ( ($ExcelContent | Where-Object { $_.Impact -ne 'Low' -and $_.'REQUIRED ACTIONS / REVIEW STATUS' -ne 'Reviewed' -and ![String]::IsNullOrEmpty($_.'REQUIRED ACTIONS / REVIEW STATUS')}).count -ge 1)
      {
        Write-Host ""
        Write-Host "There are still some recommendations that need to be reviewed." -ForegroundColor Yellow
        Write-Host "Please review all the recommendations in the Expert-Analysis file before running the report generator." -ForegroundColor Yellow
        Write-Host ""
        Exit
      }

  }

  function New-AssessmentFindingsFile {
    Param(
      [string]$AssessmentFindingsFile
      )

    $workingFolderPath = Get-Location
    $workingFolderPath = $workingFolderPath.Path
    $NewAssessmentFindings = ($workingFolderPath + '\Assessment-Findings-Report-v1-' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '.xlsx')
    # Build-WARAAssessmentFindingsCom creates this file from the template (data sheets, pivots, charts);
    # here we only compute the destination path.

    return $NewAssessmentFindings
  }

  function New-PPTFile {
    Param(
      [string]$PPTTemplateFile  # NEED FIX: PPTTemplateFile parameter does not use in the function
      )

    $workingFolderPath = Get-Location
    $workingFolderPath = $workingFolderPath.Path
    $NewPPTFile = ($workingFolderPath + '\Executive Summary Presentation - ' + $CustomerName + ' - ' + (get-date -Format "yyyy-MM-dd-HH-mm") + '.pptx')

    return $NewPPTFile
  }

  function Test-Requirement {
    # Phase 3 now reads and builds Excel via local Microsoft Excel (COM) instead of the third-party
    # ImportExcel module. Verify desktop Excel can be started via COM.
    Write-Host "Validating " -NoNewline
    Write-Host "Microsoft Excel (COM)" -ForegroundColor Cyan -NoNewline
    Write-Host " availability.."
    try {
      $probe = New-Object -ComObject Excel.Application
      $null = $probe.Version
      $probe.Quit()
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($probe)
    }
    catch {
      throw "Microsoft Excel (desktop) is required by the report generator but could not be started via COM. Install Microsoft 365 / Office desktop Excel on this machine and retry. Original error: $($_.Exception.Message)"
    }
  }

  function Set-LocalFile {
    # Define script path as the default path to save files
      $workingFolderPath = $PSScriptRoot
      Set-Location -path $workingFolderPath;
      $clonePath = "$workingFolderPath\Azure-Proactive-Resiliency-Library"
      Write-Debug "Checking the version of the script"
      $RepoVersion = Get-Content -Path "$clonePath\tools\Version.json" -ErrorAction SilentlyContinue | ConvertFrom-Json
      if ($Version -ne $RepoVersion.Generator) {
        Write-Host "This version of the script is outdated. " -BackgroundColor DarkRed
        Write-Host "Please use a more recent version of the script." -BackgroundColor DarkRed
      }
      else {
        Write-Host "This version of the script is current version. " -BackgroundColor DarkGreen
      }

  }

  function Get-ExcelImpactedResources {
    Param($ExcelFile)

    $ExcelContent = Import-WARAExcelViaCom -Path $ExcelFile -WorksheetName '4.ImpactedResourcesAnalysis' -StartRow 12
    #$ImpactedResources = $ExcelContent

    return $ExcelContent.where({![String]::IsNullOrEmpty($_."Resource Type")})
  }

  function Get-ExcelWorkloadInventory {
    Param($ExcelFile)

    $ExcelContent = Import-WARAExcelViaCom -Path $ExcelFile -WorksheetName '2.WorkloadInventory' -StartRow 12

    return $ExcelContent

  }

  function Get-ExcelPlatformIssues {
    Param($ExcelFile)

    try {
      $PlatformIssues = Import-WARAExcelViaCom -Path $ExcelFile -WorksheetName '5.PlatformIssuesAnalysis' -StartRow 12
    }
    catch {
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Platform Issues not found in the Excel File..')
    }

    return $PlatformIssues

  }

  function Get-ExcelSupportTicket {
    Param($ExcelFile)

    try {
      $SupportTickets = Import-WARAExcelViaCom -Path $ExcelFile -WorksheetName "6.SupportRequestsAnalysis" -AsText 'Ticket ID' -StartRow 12
    }
    catch {
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Support Tickets not found in the Excel File..')
    }

    return $SupportTickets
  }

  function Get-ExcelRetirement {
    Param($ExcelFile)

    try {
      $Retirements = Import-WARAExcelViaCom -Path $ExcelFile -WorksheetName "4.ImpactedResourcesAnalysis" -StartRow 12
      $Retirements = $Retirements | Where-Object {$_.Source -eq 'Azure Service Health - Service Retirements'}
    }
    catch {
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Service Retirements not found in the Excel File..')
    }

    return $Retirements

  }

  function Build-SummaryActionPlan {
    Param($ImpactedResources,$includeLow)

    if ($includeLow.IsPresent)
      {
        $Recommendations = $ImpactedResources | Where-Object {$_.impact -in ('High','Medium','Low') -and $_.ValidationCategory -ne 'Retirements'}
      }
    else
      {
        $Recommendations = $ImpactedResources | Where-Object {$_.impact -in ('High','Medium') -and $_.ValidationCategory -ne 'Retirements'}
      }

    $RecomCount = ($ImpactedResources).count
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Creating CSV file for: '+ $RecomCount +' recommendations')

    $CXSummaryArray = Foreach ($Recommendation in $Recommendations)
      {

        try{
        $tmp = [PSCustomObject]@{
          'Recommendation Guid'       = $Recommendation.Guid
          'Recommendation Title'      = $Recommendation.'Recommendation Title'
          'Description'               = $Recommendation.'Long Description'
          'Priority'                  = $Recommendation.Impact
          'Customer-facing annotation' = ""
          'Internal-facing notes'     = $Recommendation.notes
          'Potential Benefit'         = $Recommendation.'Potential Benefit'
          'Resource Type'             = $Recommendation.'Resource Type'
          'Resource ID'               = $Recommendation.id
        }
        $tmp
    } catch {
      Write-Debug $recommendation
      }
    }
    return $CXSummaryArray

  }

  ######################## PPT Functions ##########################

  ############# Slide 1
  function Remove-PPTSlide1 {
    Param($Presentation,$CustomerName,$WorkloadName)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Removing Slide 1..')

    ($Presentation.Slides | Where-Object { $_.SlideIndex -eq 1 }).Delete()

    $Slide1 = $Presentation.Slides | Where-Object { $_.SlideIndex -eq 1 }

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Slide 1 - Adding Customer name: ' + $CustomerName + '. And Workload name: ' + $WorkloadName)
    ($Slide1.Shapes | Where-Object { $_.Id -eq 5 }).TextFrame.TextRange.Text = ($CustomerName + ' - ' + $WorkloadName)
  }

  ############# SLide 12
  function Build-PPTSlide12 {
    Param($Presentation,$AUTOMESSAGE,$WorkloadName,$ResourcesTypes)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 12 - Workload Summary..')

    $SlideWorkloadSummary = $Presentation.Slides | Where-Object { $_.SlideIndex -eq 12 }

    $TargetShape = ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 9 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    $TargetShape = ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 8 })
    $TargetShape.Delete()

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 12 - Adding Workload name: ' + $WorkloadName)
    ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 3 }).TextFrame.TextRange.Text = ('During the engagement, the Workload ' + $WorkloadName + ' has been reviewed. The solution is hosted in two Azure regions, and runs mainly IaaS resources, with some PaaS resources, which includes but is not limited to:')

    $loop = 1
    foreach ($ResourcesType in $ResourcesTypes) {
      $LogResName = $ResourcesType.Name
      $LogResCount = $ResourcesType.'Count'
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 12 - Adding Resource Type: ' + $LogResName + '. Count: ' + $LogResCount)
      if ($loop -eq 1) {
        $ResourceTemp = ($ResourcesType.Name + ' (' + $ResourcesType.'Count' + ')')
        ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 6 }).Table.Columns(1).Width = 685
        ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 6 }).Table.Rows(1).Cells(1).Shape.TextFrame.TextRange.Text = $ResourceTemp
        ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 6 }).Table.Rows(1).Height = 20
      }
      else {
        $ResourceTemp = ($ResourcesType.Name + ' (' + $ResourcesType.'Count' + ')')
        ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 6 }).Table.Rows.Add() | Out-Null
        ($SlideWorkloadSummary.Shapes | Where-Object { $_.Id -eq 6 }).Table.Rows($loop).Cells(1).Shape.TextFrame.TextRange.Text = $ResourceTemp
      }
      $loop ++
    }
  }

  ############# Slide 16
  function Build-PPTSlide16 {
    Param($Presentation,$AUTOMESSAGE,$ImpactedResources)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 16 - Health and Risk Dashboard..')

    $SlideHealthAndRisk = $Presentation.Slides | Where-Object { $_.SlideIndex -eq 16 }

    $TargetShape = ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 41 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    $ServiceHighImpact = $ImpactedResources | Where-Object { $_.Impact -eq 'High' -and $_.Category -eq 'Azure Service' } | Group-Object -Property 'Recommendation Title' | Sort-Object -Property "Count" -Descending
    #$WAFHighImpact = $ImpactedResources | Where-Object { $_.Impact -eq 'High' -and $_.Category -eq 'Well Architected' } | Group-Object -Property 'Recommendation Title' | Sort-Object -Property "Count" -Descending

    $count = 1
    foreach ($Impact in $ServiceHighImpact) {
      $LogImpactName = $Impact.Name
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 16 - Adding Service High Impact Name: ' + $LogImpactName)
      if ($count -lt 7) {
          ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 9 }).TextFrame.TextRange.Paragraphs($count).text = $Impact.Name
        $count ++
      }
    }


    while (($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 9 }).TextFrame.TextRange.Paragraphs().count -gt 6) {
      ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 9 }).TextFrame.TextRange.Paragraphs(6).Delete()
    }

    <#
    if ($WAFHighImpact.count -ne 0) {
      $count = 1
      foreach ($Impact in $WAFHighImpact) {
        $LogWAFImpactName = $Impact.Name
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 16 - Adding WAF High Impact: ' + $LogWAFImpactName)
        if ($count -lt 5) {
          ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 12 }).TextFrame.TextRange.Paragraphs($count).text = $Impact.Name
          $count ++
        }
      }
    }
    else {
      ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 12 }).TextFrame.TextRange.Text = ' '
    }


    while (($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 12 }).TextFrame.TextRange.Paragraphs().count -gt 5) {
      ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 12 }).TextFrame.TextRange.Paragraphs(6).Delete()
    }
      #>

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 16 - Adding general values...')
    #Total Recomendations
    ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 19 }).TextFrame.TextRange.Text = [string]($ImpactedResources | Select-Object -Property Guid -Unique).count
    #High Impact
    ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 21 }).TextFrame.TextRange.Text = [string]($ImpactedResources | Where-Object { $_.Impact -eq 'High' } | Select-Object -Property Guid -Unique).count
    #Medium Impact
    ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 23 }).TextFrame.TextRange.Text = [string]($ImpactedResources | Where-Object { $_.Impact -eq 'Medium' } | Select-Object -Property Guid -Unique).count
    #Low Impact
    ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 24 }).TextFrame.TextRange.Text = [string]($ImpactedResources | Where-Object { $_.Impact -eq 'Low' } | Select-Object -Property Guid -Unique).count
    #Impacted Resources
    ($SlideHealthAndRisk.Shapes | Where-Object { $_.Id -eq 30 }).TextFrame.TextRange.Text = [string]($ImpactedResources | Select-Object -Property id -Unique).count
  }

    ############# Slide 17
  function Build-PPTSlide17 {
    Param($Presentation, $AUTOMESSAGE, $ExcelWorkbooks)

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 17 - Health and Risk Dashboard..')

    $ChartSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq 17 }

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 17 - Adding Automation Message..')
    $TargetShape = ($ChartSlide.Shapes | Where-Object { $_.Id -eq 41 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 17 - Loading Excel Worksheet..')
    $WS = $ExcelWorkbooks.Worksheets | Where-Object { $_.Name -eq '1.Dashboard' }
    Start-Sleep 1

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 17 - Copying Chart P1..')
    $WS.ChartObjects('ChartP1').copy()
    Start-Sleep -Milliseconds 500

    $ChartSlide.Shapes.Paste() | Out-Null
    Start-Sleep 1

    $Shape = $ChartSlide.Shapes | Where-Object { $_.Name -eq 'ChartP1' }

    $Shape.ScaleHeight(0.85, $false)
    Start-Sleep -Milliseconds 500
    $Shape.IncrementLeft(-260)
    Start-Sleep -Milliseconds 500
    $Shape.IncrementTop(77)
    Start-Sleep -Milliseconds 500

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 17 - Copying Chart P0..')
    $WS.ChartObjects('ChartP0').copy()
    Start-Sleep -Milliseconds 500

    $ChartSlide.Shapes.Paste() | Out-Null
    Start-Sleep 1

    $Shape = $ChartSlide.Shapes | Where-Object { $_.Name -eq 'ChartP0' }
    $yAxis = $Shape.Chart.Axes(1)  # 2 corresponds to xlValue for the Y-axis
    $yAxis.TickLabelSpacing = 1
    $Shape.IncrementLeft(240)
    Start-Sleep -Milliseconds 500


  }

  ############# Slide 23
  function Build-PPTSlide23 {
    Param($Presentation,$AUTOMESSAGE,$ImpactedResources)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 23 - High Impact Issues..')

    $HighImpact = $ImpactedResources | Where-Object { $_.Impact -eq 'High' } | Group-Object -Property 'Recommendation Title','Resource Type' | Sort-Object -Property "Count" -Descending

    $FirstSlide = 23
    $TableID = 6
    $CurrentSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $FirstSlide }

    $TargetShape = ($CurrentSlide.Shapes | Where-Object { $_.Id -eq 41 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 23 - Cleaning Table..')
    $row = 2
    while ($row -lt 6) {
      $cell = 1
      while ($cell -lt 5) {
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells($cell).Shape.TextFrame.TextRange.Text = ''
        $Cell ++
      }
      $row ++
    }

    $Counter = 1
    $RecomNumber = 1
    $row = 2
    foreach ($Impact in $HighImpact) {
      $LogHighImpact = $Impact.Values[0]
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 23 - Adding High Impact: ' + $LogHighImpact )
      if ($Counter -lt 14) {
        #Number
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(1).Shape.TextFrame.TextRange.Text = [string]$RecomNumber
        #Recommendation
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(2).Shape.TextFrame.TextRange.Text = $Impact.Values[0]
        #Service
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(3).Shape.TextFrame.TextRange.Text = $Impact.Values[1]
        #Impacted Resources
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(4).Shape.TextFrame.TextRange.Text = [string]$Impact.'Count'
        $counter ++
        $RecomNumber ++
        $row ++
      }
      else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 23 - Adding new Slide..')
        $Counter = 1

        $Presentation.slides[$FirstSlide].Duplicate() | Out-Null

        $FirstSlide ++

        $NextSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $FirstSlide }

        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 23 - Cleaning table of new slide..')
        $rowTemp = 2
        while ($rowTemp -lt 15) {
          $cell = 1
          while ($cell -lt 5) {
            ($NextSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($rowTemp).Cells($cell).Shape.TextFrame.TextRange.Text = ''
            $Cell ++
          }
          $rowTemp ++
        }

        $CurrentSlide = $NextSlide

        $row = 2
        #Number
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(1).Shape.TextFrame.TextRange.Text = [string]$RecomNumber
        #Recommendation
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(2).Shape.TextFrame.TextRange.Text = $Impact.Values[0]
        #Service
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(3).Shape.TextFrame.TextRange.Text = $Impact.Values[1]
        #Impacted Resources
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(4).Shape.TextFrame.TextRange.Text = [string]$Impact.'Count'
        $Counter ++
        $RecomNumber ++
        $row ++
      }
    }

  }

  ############# Slide 24
  function Build-PPTSlide24 {
    Param($Presentation,$AUTOMESSAGE,$ImpactedResources)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 24 - Medium Impact Issues..')

    $MediumImpact = $ImpactedResources | Where-Object { $_.Impact -eq 'Medium' } | Group-Object -Property 'Recommendation Title','Resource Type' | Sort-Object -Property "Count" -Descending

    $FirstSlide = 24
    $TableID = 6
    $CurrentSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $FirstSlide }

    $TargetShape = ($CurrentSlide.Shapes | Where-Object { $_.Id -eq 41 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 24 - Cleaning Table..')
    $row = 2
    while ($row -lt 6) {
      $cell = 1
      while ($cell -lt 5) {
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells($cell).Shape.TextFrame.TextRange.Text = ''
        $Cell ++
      }
      $row ++
    }

    $Counter = 1
    $RecomNumber = 1
    $row = 2
    foreach ($Impact in $MediumImpact) {
      $LogMediumImpact = $Impact.Values[0]
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 24 - Adding Medium Impact: ' + $LogMediumImpact)
      if ($Counter -lt 14) {
        #Number
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(1).Shape.TextFrame.TextRange.Text = [string]$RecomNumber
        #Recommendation
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(2).Shape.TextFrame.TextRange.Text = $Impact.Values[0]
        #Service
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(3).Shape.TextFrame.TextRange.Text = $Impact.Values[1]
        #Impacted Resources
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(4).Shape.TextFrame.TextRange.Text = [string]$Impact.'Count'
        $counter ++
        $RecomNumber ++
        $row ++
      }
      else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 24 - Creating new slide..')
        $Counter = 1

        $Presentation.slides[$FirstSlide].Duplicate() | Out-Null

        $FirstSlide ++

        $NextSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $FirstSlide }

        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 24 - Cleaning Table of new slide..')
        $rowTemp = 2
        while ($rowTemp -lt 15) {
          $cell = 1
          while ($cell -lt 5) {
            ($NextSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($rowTemp).Cells($cell).Shape.TextFrame.TextRange.Text = ''
            $Cell ++
          }
          $rowTemp ++
        }

        $CurrentSlide = $NextSlide

        $row = 2
        #Number
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(1).Shape.TextFrame.TextRange.Text = [string]$RecomNumber
        #Recommendation
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(2).Shape.TextFrame.TextRange.Text = $Impact.Values[0]
        #Service
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(3).Shape.TextFrame.TextRange.Text = $Impact.Values[1]
        #Impacted Resources
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(4).Shape.TextFrame.TextRange.Text = [string]$Impact.'Count'
        $Counter ++
        $RecomNumber ++
        $row ++
      }
    }

  }

  ############# Slide 25
  function Build-PPTSlide25 {
    Param($Presentation,$AUTOMESSAGE,$ImpactedResources)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 25 - Low Impact Issues..')

    $LowImpact = $ImpactedResources | Where-Object { $_.Impact -eq 'Low' } | Group-Object -Property 'Recommendation Title','Resource Type' | Sort-Object -Property "Count" -Descending

    $FirstSlide = 25
    $TableID = 6
    $CurrentSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $FirstSlide }

    $TargetShape = ($CurrentSlide.Shapes | Where-Object { $_.Id -eq 41 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 25 - Cleaning Table..')
    $row = 2
    while ($row -lt 6) {
      $cell = 1
      while ($cell -lt 5) {
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells($cell).Shape.TextFrame.TextRange.Text = ''
        $Cell ++
      }
      $row ++
    }

    $Counter = 1
    $RecomNumber = 1
    $row = 2
    foreach ($Impact in $LowImpact) {
      $LogLowImpact = $Impact.Values[0]
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 25 - Adding Low Impact: ' + $LogLowImpact)
      if ($Counter -lt 14) {
        #Number
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(1).Shape.TextFrame.TextRange.Text = [string]$RecomNumber
        #Recommendation
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(2).Shape.TextFrame.TextRange.Text = $Impact.Values[0]
        #Service
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(3).Shape.TextFrame.TextRange.Text = $Impact.Values[1]
        #Impacted Resources
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(4).Shape.TextFrame.TextRange.Text = [string]$Impact.'Count'
        $counter ++
        $RecomNumber ++
        $row ++
      }
      else {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 25 - Creating new Slide..')
        $Counter = 1

        $Presentation.slides[$FirstSlide].Duplicate() | Out-Null

        $FirstSlide ++

        $NextSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $FirstSlide }

        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 25 - Cleaning Table of new slide..')
        $rowTemp = 2
        while ($rowTemp -lt 15) {
          $cell = 1
          while ($cell -lt 5) {
            ($NextSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($rowTemp).Cells($cell).Shape.TextFrame.TextRange.Text = ''
            $Cell ++
          }
          $rowTemp ++
        }

        $CurrentSlide = $NextSlide

        $row = 2
        #Number
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(1).Shape.TextFrame.TextRange.Text = [string]$RecomNumber
        #Recommendation
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(2).Shape.TextFrame.TextRange.Text = $Impact.Values[0]
        #Service
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(3).Shape.TextFrame.TextRange.Text = $Impact.Values[1]
        #Impacted Resources
        ($CurrentSlide.Shapes | Where-Object { $_.Id -eq $TableID }).Table.Rows($row).Cells(4).Shape.TextFrame.TextRange.Text = [string]$Impact.'Count'
        $Counter ++
        $RecomNumber ++
        $row ++
      }
    }

  }

  ############# Slide 28
  function Build-PPTSlide28 {
    Param($Presentation,$AUTOMESSAGE,$PlatformIssues)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 28 - Recent Microsoft Outages..')

    $Loop = 1
    $CurrentSlide = 28

    if (![string]::IsNullOrEmpty($PlatformIssues)) {
      Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 28 - Outages found..')
      foreach ($Outage in $PlatformIssues) {
        if (![string]::IsNullOrEmpty($Outage.title))
            {
                if ($Loop -eq 1) {
                    $OutageName = ($Outage.'Tracking ID' + ' - ' + $Outage.title)
                    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 28 - Adding Outage: ' + $OutageName)

                    $OutageService = $Outage.'Impacted Service'

                    $SlidePlatformIssues = $Presentation.Slides | Where-Object { $_.SlideIndex -eq 28 }

                    $TargetShape = ($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 4 })
                    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

                    ($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = $OutageName

                    ($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(3).Text = $Outage.'What happened'

                    ($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(5).Text = $OutageService

                    ($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(7).Text = $Outage.'How can customers make incidents like this less impactful'

                    while (($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs().count -gt 7) {
                      ($SlidePlatformIssues.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(8).Delete()
                    }
                  }
                  else {
                    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 28 - Creating new Slide..')
                    ############### NEXT 9 SLIDES

                    $OutageName = ($Outage.'Tracking ID' + ' - ' + $Outage.title)
                    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 28 - Adding Outage: ' + $OutageName)

                    $OutageService = $Outage.'Impacted Service'

                    $Presentation.slides[$CurrentSlide].Duplicate() | Out-Null

                    $CurrentSlide ++

                    $NextSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $CurrentSlide }

                    ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = $OutageName

                    ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(3).Text = $Outage.'What happened'

                    ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(5).Text = $OutageService

                    ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(7).Text = $Outage.'How can customers make incidents like this less impactful'

                    while (($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs().count -gt 7) {
                      ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(8).Delete()
                    }
                  }
                  $Loop ++
            }
      }
    }
  }

  ############# Slide 29
  function Build-PPTSlide29 {
    Param($Presentation,$AUTOMESSAGE,$SupportTickets)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 29 - Sev-A Support Requests..')

    $Loop = 1
    $CurrentSlide = 29
    $Slide = 1

    # Prevents the script from creating the SupportTickets slide if there are no support tickets
    if(!($SupportTickets.count -eq 1 -and [string]::IsNullOrEmpty($SupportTickets.'Ticket ID')))
    {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 29 - Support Tickets found..')
        foreach ($Tickets in $SupportTickets) {
            $TicketName = ($Tickets.'Ticket ID' + ' - ' + $Tickets.Title)
            Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 29 - Adding Ticket: ' + $TicketName)
            $TicketStatus = $Tickets.'Status'
            $TicketDate = $Tickets.'Creation Date'

        if ($Slide -eq 1) {
            if ($Loop -eq 1) {
                $SlideSupportRequests = $Presentation.Slides | Where-Object { $_.SlideIndex -eq 29 }
                $TargetShape = ($SlideSupportRequests.Shapes | Where-Object { $_.Id -eq 4 })
                $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

                ($SlideSupportRequests.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = $TicketName
                ($SlideSupportRequests.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(2).Text = "Status: $TicketStatus"
                ($SlideSupportRequests.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(3).Text = "Creation Date: $TicketDate"

                $Loop ++
            }
            else {

                $newtextbox = ($SlideSupportRequests.Shapes | Where-Object { $_.Id -eq 7 }).Duplicate()

                if ($Loop -eq 2) {
                    $newtextbox.name = 'Shape2'
                    $newtextbox.top = [int]$newtextbox.top + 100
                    $newtextbox.left = [int]$newtextbox.left - 10
                }
                elseif ($Loop -eq 3) {
                    $newtextbox.name = 'Shape3'
                    $newtextbox.top = [int]$newtextbox.top + 220
                    $newtextbox.left = [int]$newtextbox.left - 10
                }

                $newtextbox.TextFrame.TextRange.Paragraphs(1).Text = $TicketName
                $newtextbox.TextFrame.TextRange.Paragraphs(2).Text = "Status: $TicketStatus"
                $newtextbox.TextFrame.TextRange.Paragraphs(3).Text = "Creation Date: $TicketDate"

                if ($Loop -eq 3) {
                    $Loop = 1
                    $Slide ++
                }
                else {
                    $Loop ++
                }
                Start-Sleep -Milliseconds 500
            }
            }
            else {
            if ($Loop -eq 1) {
                Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 29 - Adding new Slide..')

                $Presentation.slides[$CurrentSlide].Duplicate() | Out-Null

                $CurrentSlide ++

                $NextSlide = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $CurrentSlide }

                ($NextSlide.Shapes | Where-Object { $_.name -eq 'Shape2' }).Delete()
                ($NextSlide.Shapes | Where-Object { $_.name -eq 'Shape3' }).Delete()

                ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = $TicketName
                ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(2).Text = "Status: $TicketStatus"
                ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(3).Text = "Creation Date: $TicketDate"

                $Loop ++
            }
            else {
                $newtextbox = ($NextSlide.Shapes | Where-Object { $_.Id -eq 7 }).Duplicate()

                if ($Loop -eq 2) {
                    $newtextbox.name = 'Shape2'
                    $newtextbox.top = [int]$newtextbox.top + 100
                    $newtextbox.left = [int]$newtextbox.left - 10
                }
                elseif ($Loop -eq 3) {
                    $newtextbox.name = 'Shape3'
                    $newtextbox.top = [int]$newtextbox.top + 220
                    $newtextbox.left = [int]$newtextbox.left - 10
                }

                $newtextbox.TextFrame.TextRange.Paragraphs(1).Text = $TicketName
                $newtextbox.TextFrame.TextRange.Paragraphs(2).Text = "Status: $TicketStatus"
                $newtextbox.TextFrame.TextRange.Paragraphs(3).Text = "Creation Date: $TicketDate"

                if ($Loop -eq 3) {
                    $Loop = 1
                    $Slide ++
                }
                else {
                    $Loop ++
                }
            }
            }
            Start-Sleep -Milliseconds 500
        }
    }
    else
    {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 29 - No Support Tickets found..')
        ($Presentation.Slides | Where-Object { $_.SlideIndex -eq 29 }).Delete()
    }
  }

  ############# Slide 30
  function Build-PPTSlide30 {
    Param($Presentation,$AUTOMESSAGE,$Retirements)
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 30 - Service Retirement Notifications..')

    $Loop = 1
    $Spacer = 1
    $CurrentSlide = 30

    if (![string]::IsNullOrEmpty($Retirements)) {
    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 30 - Service Retirement found..')
    $SlideRetirements = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $CurrentSlide }

    $TargetShape = ($SlideRetirements.Shapes | Where-Object { $_.Id -eq 4 })
    $TargetShape.TextFrame.TextRange.Text = $AUTOMESSAGE

    ($SlideRetirements.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = '.'

    foreach ($Retirement in $Retirements) {

        $RetireName = ($Retirement.'Retirement TrackingId' + ' - ' + $Retirement.'Recommendation Title')
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 30 - Adding Retirement: ' + $RetireName)
        if ($Loop -lt 11) {
            if ($Loop -eq 1) {

                ($SlideRetirements.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = $RetireName
                $Loop ++
            }
            else {

                $NewShape = ($SlideRetirements.Shapes | Where-Object { $_.Id -eq 7 }).Duplicate()

                $NewShape.name = ('Shape_'+$Loop)

                $NewShape.top = [int]$NewShape.top + ($Spacer * 40)
                $NewShape.left = [int]$NewShape.left - 10

                $NewShape.TextFrame.TextRange.Paragraphs(1).Text = $RetireName

                $Loop ++
                $Spacer ++
            }
        }
        else {
            Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Editing Slide 30 - Creating new Slide..')
            $Loop = 1
            $Spacer = 1

            $Presentation.slides[$CurrentSlide].Duplicate() | Out-Null

            $CurrentSlide ++

            $SlideRetirements = $Presentation.Slides | Where-Object { $_.SlideIndex -eq $CurrentSlide }

            $ShapesToDelete = $SlideRetirements.Shapes | Where-Object { $_.Name -like 'Shape_*'}

            Foreach ($Shape in $ShapesToDelete)
                {
                    $Shape.Delete()
                }

            ($SlideRetirements.Shapes | Where-Object { $_.Id -eq 7 }).TextFrame.TextRange.Paragraphs(1).Text = $RetireName
            $Loop ++
            }
        }
    }
  }

  ######################## Assessment Findings Functions ##########################

  function Initialize-ExcelImpactedResources {
    Param($ImpactedResources)

    $ImpactedResourcesFormatted = @()

    ForEach ($Resource in $ImpactedResources) {
      $obj = @{
        'Impacted?' = 'Yes';
        'Resource Type' = $Resource.'Resource Type';
        'subscriptionId' = $Resource.subscriptionId;
        'resourceGroup' = $Resource.resourceGroup;
        'location' = $Resource.location;
        'name' = $Resource.name;
        'id' = $Resource.id;
        'custom1' = $Resource.custom1;
        'custom2' = $Resource.custom2;
        'custom3' = $Resource.custom3;
        'custom4' = $Resource.custom4;
        'custom5' = $Resource.custom5;
        'Recommendation Title' = $Resource.'Recommendation Title';
        'Impact' = $Resource.Impact;
        'Recommendation Control' = $Resource.'Recommendation Control';
        'Potential Benefit' = $Resource.'Potential Benefit';
        'Learn More Link' = $Resource.'Learn More Link';
        'Long Description' = $Resource.'Long Description';
        'Guid' = $Resource.Guid;
        'Category' = $Resource.Category;
        'Source' = $Resource.'Source';
        'WAF Pillar' = $Resource.'WAF Pillar';
        'Platform Issue TrackingId' = $Resource.'Platform Issue TrackingId';
        'Retirement TrackingId' = $Resource.'Retirement TrackingId';
        'Support Request Number' = $Resource.'Support Request Number';
        'Notes' = $Resource.Notes;
        'checkName' = $Resource.checkName
      }
      $ImpactedResourcesFormatted += $obj
    }

    # Returns the array with all the recommendations already formatted to be exported to Excel
    return $ImpactedResourcesFormatted

  }

  # NOTE: Export-ExcelImpactedResources removed — the '3.ImpactedResources' sheet is written via
  # Excel COM in Build-WARAAssessmentFindingsCom (reports/WARAReportExcelCom.ps1) — no ImportExcel.

  function Initialize-ExcelRecommendations {
    Param($ImpactedResources)

    $RecommendationsFormatted = @()

    $GroupedResources = $ImpactedResources.where({![String]::IsNullOrEmpty($_.Guid)}) | Group-Object -Property 'Guid' | Sort-Object -Property 'Count' -Descending

    $CustomRecommendations = $ImpactedResources.where({[String]::IsNullOrEmpty($_.Guid)})

    ForEach ($Resource in $GroupedResources) {

      $Recommendation = $ImpactedResources | Where-Object { $_.Guid -eq $Resource.Name } | Select-Object -First 1

      $obj = @{
        'Impact' = $Recommendation.Impact;
        'Description' = $Recommendation.'Recommendation Title';
        'Potential Benefit' = $Recommendation.'Potential Benefit';
        'Impacted Resources' = $Resource.Count;
        'Resource Type' = $Recommendation.'Resource Type';
        'Recommendation Control' = $Recommendation.'Recommendation Control';
        'Long Description' = $Recommendation.'Long Description';
        'Category' = $Recommendation.Category;
        'Learn More Link' = $Recommendation.'Learn More Link';
        'Guid' = $Recommendation.Guid;
        'Notes' = $Recommendation.Notes;
      }
      $RecommendationsFormatted += $obj
    }

    $CustomRecommendations.Foreach(
        {
            $obj = @{
                'Impact' = $_.Impact;
                'Description' = $_.'Recommendation Title'
                'Potential Benefit' = $_.'Potential Benefit';
                'Impacted Resources' = 1;
                'Resource Type' = $_.'Resource Type';
                'Recommendation Control' = $_.'Recommendation Control';
                'Long Description' = $_.'Long Description';
                'Category' = $_.Category;
                'Learn More Link' = $_.'Learn More Link';
                'Guid' = $_.Guid;
                'Notes' = $_.Notes;
            }
            $RecommendationsFormatted += $obj
        }
    )


    # Returns the array with all the recommendations already formatted to be exported to Excel
    return $RecommendationsFormatted

  }

  # NOTE: Export-ExcelRecommendations, Export-ExcelWorkloadInventory, Build-ExcelPivotTable and
  # Build-ExcelPivotChart were removed. The Assessment-Findings workbook (data sheets + tables), its
  # pivot tables (P0-P3) and its charts (ChartP0/ChartP1) are now built with local Microsoft Excel
  # via COM in Build-WARAAssessmentFindingsCom (reports/WARAReportExcelCom.ps1) — no ImportExcel.

# Start the stopwatch to time the script
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

#Call the functions
$Version = "2.1.5"
Write-Host "Version: " -NoNewline
Write-Host $Version -ForegroundColor DarkBlue -NoNewline
Write-Host " "

$ExpertAnalysisSourceFile = (Get-Item -Path $ExpertAnalysisFile).FullName

# Excel COM cannot reliably open a workbook that lives on an actively-syncing OneDrive path (the open is
# rejected with RPC_E_CALL_REJECTED / forced into Protected View, which aborts the read). Work from a
# local temp copy of the Expert-Analysis workbook and read everything from there.
$CoreFile = Join-Path $env:TEMP ('WARA-ExpertAnalysis-' + [guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item -LiteralPath $ExpertAnalysisSourceFile -Destination $CoreFile -Force
try { Unblock-File -LiteralPath $CoreFile -ErrorAction SilentlyContinue } catch {}

# Read the Expert-Analysis worksheets once via Excel COM (replaces ImportExcel; no third-party module).
Initialize-WARAExpertAnalysisCache -Path $CoreFile -Sheets @(
    @{ Name = '4.ImpactedResourcesAnalysis'; StartRow = 12 },
    @{ Name = '2.WorkloadInventory';         StartRow = 12 },
    @{ Name = '5.PlatformIssuesAnalysis';    StartRow = 12 },
    @{ Name = '6.SupportRequestsAnalysis';   StartRow = 12; AsText = @('Ticket ID') }
)

Test-ReviewedRecommendations -ExcelFile $CoreFile

Write-Debug (' ---------------------------------- STARTING REPORT GENERATOR SCRIPT --------------------------------------- ')
Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Starting Report Generator Script..')
Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Script Version: ' + $Version)
Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Excel File: ' + $ExpertAnalysisFile)

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "10% Complete." -PercentComplete 10
Test-Requirement
Write-Progress -Id 1 -activity "Processing Office Apps" -Status "15% Complete." -PercentComplete 15
#Set-LocalFile
Write-Progress -Id 1 -activity "Processing Office Apps" -Status "20% Complete." -PercentComplete 20

$ExcelImpactedResources = Get-ExcelImpactedResources -ExcelFile $CoreFile

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "25% Complete." -PercentComplete 25

$ExcelPlatformIssues = Get-ExcelPlatformIssues -ExcelFile $CoreFile

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "30% Complete." -PercentComplete 30

$ExcelSupportTickets = Get-ExcelSupportTicket -ExcelFile $CoreFile

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "35% Complete." -PercentComplete 35

$ExcelWorkloadInventory = Get-ExcelWorkloadInventory -ExcelFile $CoreFile

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "40% Complete." -PercentComplete 40

$ExcelRetirements = Get-ExcelRetirement -ExcelFile $CoreFile

# Done reading the Expert-Analysis workbook; remove the local temp copy.
Remove-Item -LiteralPath $CoreFile -Force -ErrorAction SilentlyContinue

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "45% Complete." -PercentComplete 45


$PPTFinalFile = New-PPTFile -PPTTemplateFile $pptTemplateFilePath  # NEED FIX: PPTTemplateFile parameter does not use in the function
Write-Host "PowerPoint" -ForegroundColor DarkRed -NoNewline
Write-Host " and " -NoNewline
Write-Host "Excel" -ForegroundColor DarkGreen -NoNewline
Write-Host " "
Write-Host "Editing " -NoNewline
$NewAssessmentFindingsFile = New-AssessmentFindingsFile -AssessmentFindingsFile $AssessmentFindingsFile


$AUTOMESSAGE = 'AUTOMATICALLY MODIFIED (Please Review)'

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Resource Types..')

$TempImpactedResources = $ExcelImpactedResources | Select-Object -Property 'Resource Type', 'id' -Unique

$ResourcesTypes = $TempImpactedResources | Group-Object -Property 'Resource Type' | Sort-Object -Property 'Count' -Descending | Select-Object -First 10

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Starting to Process Assessment Findings..')

$ImpactedResourcesFormatted = Initialize-ExcelImpactedResources -ImpactedResources $ExcelImpactedResources
$RecommendationsFormatted = Initialize-ExcelRecommendations -ImpactedResources $ExcelImpactedResources

# Build the entire Assessment-Findings workbook (data sheets + tables + pivots + charts) with Excel COM.
$null = Build-WARAAssessmentFindingsCom `
    -TemplatePath $AssessmentFindingsFile `
    -OutputPath $NewAssessmentFindingsFile `
    -ImpactedResources @($ImpactedResourcesFormatted | ForEach-Object { [PSCustomObject]$_ }) `
    -Recommendations @($RecommendationsFormatted | ForEach-Object { [PSCustomObject]$_ }) `
    -WorkloadInventory @($ExcelWorkloadInventory | ForEach-Object { [PSCustomObject]$_ })

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Starting PowerPoint..')
# Openning PPT
try
    {
        $Application = New-Object -ComObject PowerPoint.Application
        $Presentation = $Application.Presentations.Open($pptTemplateFilePath, $null, $null, $null)
    }
catch
    {
        Write-Host ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + "- Error: " + $_.Exception.Message)
        Get-Process -Name "POWERPNT" -ErrorAction Ignore | Where-Object { $_.CommandLine -like '*/automation*' } | Stop-Process -Force
    }

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Starting Excel..')
# Openning Excel
try
    {
        $ExcelApplication = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 500
        $ExcelWorkbooks = $ExcelApplication.Workbooks.Open($NewAssessmentFindingsFile)
    }
catch
    {
        Write-Host ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + "- Error: " + $_.Exception.Message)
        Get-Process -Name "excel" -ErrorAction Ignore | Where-Object { $_.CommandLine -like '*/automation*' } | Stop-Process -Force
    }


Remove-PPTSlide1 -Presentation $Presentation -CustomerName $CustomerName -WorkloadName $WorkloadName
Build-PPTSlide12 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -WorkloadName $WorkloadName -ResourcesType $ResourcesTypes
Build-PPTSlide16 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -ImpactedResources $ExcelImpactedResources

while ([string]::IsNullOrEmpty($ExcelWorkbooks)) {
    Start-Sleep 1
}

Build-PPTSlide17 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -ExcelWorkbooks $ExcelWorkbooks

Build-PPTSlide30 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -Retirements $ExcelRetirements

Build-PPTSlide29 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -SupportTickets $ExcelSupportTickets

Build-PPTSlide28 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -PlatformIssues $ExcelPlatformIssues

Build-PPTSlide25 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -ImpactedResources $ExcelImpactedResources
Build-PPTSlide24 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -ImpactedResources $ExcelImpactedResources
Build-PPTSlide23 -Presentation $Presentation -AUTOMESSAGE $AUTOMESSAGE -ImpactedResources $ExcelImpactedResources

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Closing Excel..')
$ExcelWorkbooks.Save()
$ExcelWorkbooks.Close()
$ExcelApplication.Quit()

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Closing PowerPoint..')
$Presentation.SaveAs($PPTFinalFile)
$Presentation.Close()
$Application.Quit()

#if ($csvExport.IsPresent) {
    $WorkloadRecommendationTemplate = Build-SummaryActionPlan -ImpactedResources $ExcelImpactedResources -includeLow $includeLow

    $CSVFile = ("$PWD\Impacted Resources and Recommendations Template " + (get-date -Format "yyyy-MM-dd-HH-mm") + '.csv')

    $WorkloadRecommendationTemplate | Export-Csv -Path $CSVFile
#}

if (Get-Process -Name "POWERPNT" -ErrorAction Ignore | Where-Object { $_.CommandLine -like '*/automation*' })
    {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Trying to kill PowerPoint process.')
        Get-Process -Name "POWERPNT" -ErrorAction Ignore | Where-Object { $_.CommandLine -like '*/automation*' } | Stop-Process -Force
    }

if (Get-Process -Name "excel" -ErrorAction Ignore | Where-Object { $_.CommandLine -like '*/automation*' } )
    {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Trying to kill Excel process.')
        Get-Process -Name "excel" -ErrorAction Ignore | Where-Object { $_.CommandLine -like '*/automation*' } | Stop-Process -Force
    }

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "90% Complete." -PercentComplete 90

Write-Progress -Id 1 -activity "Processing Office Apps" -Status "100% Complete." -Completed

$stopwatch.Stop()

################ Finishing

Write-Host "---------------------------------------------------------------------"
Write-Host ('Execution Complete. Total Runtime was: ') -NoNewline
Write-Host $stopwatch.Elapsed.toString('hh\:mm\:ss') -ForegroundColor Cyan
Write-Host 'PowerPoint File Saved as: ' -NoNewline
Write-Host $PPTFinalFile -ForegroundColor Cyan
Write-Host 'Assessment Findings File Saved as: ' -NoNewline
Write-Host $NewAssessmentFindingsFile -ForegroundColor Cyan

#if ($csvExport.IsPresent) {
    Write-Host 'CSV File Saved as: ' -NoNewline
    Write-Host $CSVFile -ForegroundColor Cyan
#}

Write-Host "---------------------------------------------------------------------"

# SIG # Begin signature block
# MIIoLAYJKoZIhvcNAQcCoIIoHTCCKBkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAtwQlTMMRy8CB1
# tQzKzxcvI8zsYUQJTW+JAY1GDBgwhqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIN2w0+ubM5FX1+11xDAcqIDq
# QjhbwCvwC/15WZFN6BTnMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQA67rob4DQkuL2UhPMG9CqrMMa8fh1GY3yaXpFgGhHn6hFOEuUiqvo9
# zwAeO3ZJ6xuHhet3BZy/4aAF7ggN1zGHs+bNZIb5EqQlF52DKLm2ioehj/VUVQ6+
# c9lvGuTPz2UUlKYp6VAj7L3F/xjVVcyCTGVM1nS5jslSe0Wc4Cq0l21SBG+gFyY6
# 7wkzfc2yB7AvImwp8/56Qd7fqfxZ/1CY2NpcULz2b5XnczAaIoUtvTyWJ34qjtGD
# s+TDa/0BbshUY2sE2UjQxeG3bNxNgRfI74YlBAPRkerwxT6go5J+5whNWLuVWzbq
# 8HySU5bfi7SYj8yO8MkfVJvLdxAoGeEnoYIXlDCCF5AGCisGAQQBgjcDAwExgheA
# MIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIGor4jb9Y6dHnMGcl6Hb2Tmgs1gcQum/CkF0SSqAy9RhAgZoJmwI
# D34YEzIwMjUwNTIyMTQ1OTExLjczMlowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpFMDAy
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEeowggcgMIIFCKADAgECAhMzAAACCxGdVimS+b+FAAEAAAILMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI1OFoXDTI2MDQyMjE5NDI1OFowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpFMDAyLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAKqs+K1GMBeoWHYgfNBgPe7KQM/IkX7v17kKRjhp
# ixv3dBW6fh9ncm7/RYeu0AXcmzBKIBWXunUfbMUqu/KFpW5FhxJzti7I/QYcRWrt
# hVLm4XdguUt6jIY48pJcazpqBbrQKpUGE9D7afj17pgP3Aa/pjQeYEqadCpj8Aom
# KOBSJG6VvWi80RxNmZM1G7WZj7QmOIS6z54XrpU2gG/VrYxbQk9hsUGE/MUeV10V
# dyN/aoiRvkFhawUvtCUGNcCl/YeY2s/MbfqDXJuj8qdGNKAhGkd5hfEEhoIQa1HU
# Sakn3Q5IUfFS7t4iNV96eqdW1qDTIBBjvMRZJcw3r4IeO/dIE22blkhHLIKxRkbz
# jngr7zrlvDsXC4fy+TlD7TsrmUVjTZ4EPDdQrNaNa+pOnakrwxNriSjO+UqCtIan
# myHnsaeSOW82+3vw1dtlNUivBLvFwgaNu9L/avmVENP5Dc/a1P6DooM90ue2VxsV
# 0e67PInLHRZ8KvzMe/zAYIHeo85tFcriikAidJDiPmJLbwoBMkbBBptG55m+G930
# 53TBwVg5viitS/V0PlfAXIqUYM8xxM85CvSDGYRBWsuIkSgvLJbwRX9oMVRdOjHK
# F/NZnoNzSX1562HUS3Y1DcD+oFHnHfemMp0clUnnGmnYAjdnM1gsMKo7WW2HPUZp
# LP49AgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUZlwsiN+v6XpZ4y6ET2CzF73yP/4w
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAJoo9WfoYtyeRs8R0zRH2rhmbT7sl8bT
# n6mao/qcDhLBPbqT0eUlCOrD/5aSJAsPijAnGigiv0wNluJlfrKk+EX6DU/m7wE4
# Rc9hiQ5anMUYtEu2Vke9/GmoAhruaMsYcPMqv/XTYBdoPiYbJ0TBosaFSf0MEZ0y
# bl2GKKxYCpv8s3iWbg65fhguopxb/TJJwcqCmsaAtj7qmq/TM9TviPt4Kv48NM1g
# zy0Un9u1ScE0vI4ThZnAEiri2e9eN3ZwI62BSehCMVTVljtlZ2fLcElt/eEGtvc0
# HYhTie7rba1WtvV82TB29PnTVZFQEx32gV/jGJ+PaPYKVu/VcaWNeZ3rImO+35sG
# 68ktS5z2b68wo5bJHBR286X8TlrzNXcyuMKdjp8istA3ME5nqSUgI8KL7YaptWIz
# QIjoJtnKl6LlGv2ElsfytkTONphYuev9c+xTTCvZOfov10nHar8sCgr6i+IAiePR
# o/iJGszJRI7oovYrBjrMlzpIPFseTBgd1BGs80/dK60yDouAFK+2Z+rCBpcgeQKi
# JQUQDP1wYiq+CMakamkpNxO6ijcQSCn/NF3TDFbRCqZXfKBgaJx6ffQoBYesoULx
# JvBzqoLimh1VlglmaKscE/+6juxCd7kfTG128OZcNR/LmA2MuWNAhSwmFBLLJdtK
# NJvZrmOE5R/HMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046RTAwMi0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AKhCd1Qk3c+mrbGHxGG2xHfk0JnBoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2XvJMCIYDzIwMjUwNTIyMTAz
# MjQxWhgPMjAyNTA1MjMxMDMyNDFaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOvZ
# e8kCAQAwBwIBAAICOt8wBwIBAAICEiUwCgIFAOvazUkCAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkq
# hkiG9w0BAQsFAAOCAQEAjLBRDG7uJ+9VEuvhSBw+DxuQZSrOhkpFjpzVXZ9aYVwu
# LyqcQXnT52lWM8wqaAcBH3AVS1T0l9gPtv8ZO+F+l644NSsUizTvn93xYouEBs6Z
# Jqd3EbKMKmp7RYKkcThJ1omn1kXKyAhaVTk4gc7IVyHPobcsBXpK33Jl/WF0MrlP
# SzFB7J0XusHzHsMWyH0rr7NL4yDsuG8yRNQv7O7sCFnmGVp+ePSHKmWkLUoFPhFV
# MlOz4uQxypQFYbVBdocsz7/ufW4SAj1pgZ12hT8v5LHpE4SHQb1uK6rUMZtTBlZL
# YFIFQ3cNoXXAPwkEvY6f7z+O2gwPVD4WWKyRJBJdczGCBA0wggQJAgEBMIGTMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACCxGdVimS+b+FAAEAAAIL
# MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# LwYJKoZIhvcNAQkEMSIEIGqEyq93vnuHQaHAlVtW1JAl6sDaPh0QYzmSAF+6pFWX
# MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgNNV0q7o3vtuHTB07IX6iBE3y
# 8olzDmOd/b8S6yxiSjMwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAgsRnVYpkvm/hQABAAACCzAiBCAEP1GiUtIVayJzIES1PXhTo9uD
# RpXC0PropG2gldwFZzANBgkqhkiG9w0BAQsFAASCAgAMXnYZOjz1n/Nyfqmbe52a
# 5WG+9Lkw3pLrOD9I1MTfUMBuopzsHzEruHwfsS4U1y+DJsFOK0OkbumbAQJhZVR7
# qb5FQfTPb8k8r8hGs45j8DL8TPsIM8X1tR/Mu8F6kscZpNuAL3oMoPfpcqzrmYGB
# H0UMEfcj3VCXvhMM2Q1A7PMqu7mZmOrFMqnS45FzsYHjHfr5LhI7HFRWOkBPXC82
# E3eM53bwznCjO1zKEIYNuClfpUZxF01amRKRBwq6f/FZP+qyHniuS3XsjELZMZ7v
# KfVKgoZqYn7bmercXTZoxwg+EZy30RXPRwMDtYY7BoNY4BVmCcZY2E7/4ZelJwgA
# v4JPUQ0BIbovFVXiHdYk0E13izkXs678OgtU82i9d4LwssxfLcCt32AOawEYlAzr
# BQzttNRFGRomPbZ58YFcZRv3XR/dhZZS4JmIDxnetqmX2un4+KZuBXkZ/qCKmNQ7
# 936NnPHZFnt5xXY4cPbF3v0GH0oo7BVUwsEA2Wtoqd4V42/Wt+iFosLYKV+7JUUX
# 5MARGD0wv7/jdk6OuJItOG0LHTNymGqGGmQmTkx+DY+7clBIr3a219neDREqeFNs
# n4sZKaoDj0B4u0vMc+fTZXE1IWK2IrZbIlegNUdiTeygwNn76LHxwnBbA+Yc4/ze
# qWn4L68D41fBd0MUKPi0EQ==
# SIG # End signature block
