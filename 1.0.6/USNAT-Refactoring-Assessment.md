# WARA Module — USNAT / Air‑Gapped Sovereign Cloud Refactoring Assessment

**Module:** `wara` v1.0.6
**Target environment:** USNAT (US National / sovereign Azure cloud) — different ARM/AAD/Graph endpoints, **no public internet connectivity**.
**Goal:** Make every code path work when (a) the host cannot reach the public internet (github.com, azure.github.io, PowerShell Gallery), and (b) all Azure endpoints are the USNAT endpoints, discovered from `Get-AzContext` rather than hard‑coded to the commercial cloud.

---

## 1. Executive summary

The module has **four classes of external dependency**. Only two of them are genuine blockers for USNAT, and both are fixable:

| # | Dependency class | Reachable in USNAT? | Fixable? | Effort |
|---|------------------|---------------------|----------|--------|
| A | **Azure control‑plane calls** (`Search-AzGraph`, `Invoke-AzRestMethod`, `Get-AzAccessToken`, `Connect-AzAccount`) | ✅ Yes — these already honor the `AzContext` environment | Already OK, minor hardening needed | Low |
| B | **Public content downloads** (`recommendations.json`, `WARAinScopeResTypes.csv` from `azure.github.io` / `raw.githubusercontent.com`) | ❌ No — public GitHub Pages | ✅ Bundle the files locally (done — see `offline-data/`) | Medium |
| C | **PowerShell Gallery calls** (`Find-Module` version check, `Install-Module ImportExcel`) | ❌ No (unless an internal repo is registered) | ✅ Make optional / point to internal repo / pre‑stage modules | Low‑Medium |
| D | **Hard‑coded commercial endpoints** (`https://management.azure.com/` default, `AzureEnvironment = 'AzureCloud'` default) | ❌ Wrong host for USNAT | ✅ Derive from `Get-AzContext` | Low |

**The single most important architectural fact:** every *authenticated Azure* call in this module (Resource Graph queries and REST calls) is already routed through the `Az.Accounts` context. `Invoke-AzRestMethod` and `Search-AzGraph` build their host from `(Get-AzContext).Environment`, so **once you `Connect-AzAccount -Environment <USNAT>` they automatically target the correct sovereign endpoints.** The only place a raw host is hard‑coded is the Advisor‑metadata default parameter (Class D), and the collector already overrides it with the context value.

The genuine work is therefore: **(1) bundle the two public data files, (2) stop calling the PowerShell Gallery, and (3) fix the commercial‑cloud defaults so the environment is taken from `Get-AzContext`.**

---

## 2. Complete inventory of outbound calls

Locations are file + line as of v1.0.6.

### 2.1 Azure control‑plane (Class A — already environment‑aware)

| Call | File / Line | Notes |
|------|-------------|-------|
| `Search-AzGraph … -UseTenantScope` / `-Subscription` | [utils/utils.psm1](utils/utils.psm1#L47) and [utils/utils.psm1](utils/utils.psm1#L55) | Az.ResourceGraph uses the context's ARM endpoint. Works in USNAT **if** the ARG service is enabled in that cloud (see Risk R1). |
| `Invoke-AzRestMethod` (via `Invoke-AzureRestApi` → `Get-AzureRestMethodUriPath`) | [utils/utils.psm1](utils/utils.psm1#L173), path builder [utils/utils.psm1](utils/utils.psm1#L252) | Builds a **relative** path (`/subscriptions/…`); host is supplied by the context. ✅ Sovereign‑safe. |
| `Get-AzAccessToken -ResourceUrl $ResourceURL` | [advisor/advisor.psm1](advisor/advisor.psm1#L296) | `$ResourceURL` is passed the context `ResourceManagerUrl` by the collector. ✅ when called from the collector; ⚠️ default is commercial (Class D). |
| `Connect-AzAccount -Environment $AzureEnvironment` | [utils/utils.psm1](utils/utils.psm1#L409) | Correct mechanism, but default value is `AzureCloud` (Class D). |
| Resource‑Health REST (`Microsoft.ResourceHealth/events`) for outages/retirements | [outage/outage.psm1](outage/outage.psm1#L64), [retirement/retirement.psm1](retirement/retirement.psm1#L57) | Via `Invoke-AzureRestApi`. ✅ environment‑aware. Data availability is a risk (R2). |
| Advisor / Service Health / Support ARG queries | [advisor/advisor.psm1](advisor/advisor.psm1), [servicehealth/servicehealth.psm1](servicehealth/servicehealth.psm1#L40), [support/support.psm1](support/support.psm1#L52) | Via `Invoke-WAFQuery` → `Search-AzGraph`. ✅ environment‑aware. |

### 2.2 Public internet content (Class B — MUST bundle)

| Call | File / Line | URL |
|------|-------------|-----|
| `Invoke-RestMethod $RecommendationDataUri` | [wara.psm1](wara.psm1#L267) | `https://azure.github.io/WARA-Build/objects/recommendations.json` |
| `Invoke-RestMethod $RecommendationResourceTypesUri` | [wara.psm1](wara.psm1#L280) | `https://azure.github.io/WARA-Build/objects/WARAinScopeResTypes.csv` |
| `Invoke-RestMethod $RecommendationDataUri` (analyzer) | [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L228) | same JSON |
| `Invoke-RestMethod $RecommendationResourceTypesUri` (analyzer) | [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L1165) | same CSV |
| Default URIs in params | [wara.psm1](wara.psm1#L146), [wara.psm1](wara.psm1#L152), [analyzer/analyzer.psm1](analyzer/analyzer.psm1#L27), [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L31) | GitHub Pages |

> ✅ **Both files have already been downloaded to [offline-data/](offline-data/)** (`recommendations.json` — 393 records; `WARAinScopeResTypes.csv` — 3,861 rows) so they can be bundled with the package or hosted on an internal USNAT web server / storage account.
>
> ℹ️ **Note on the data content:** `recommendations.json` contains many `"url"` fields pointing at public documentation (e.g. `learn.microsoft.com`, `kubernetes.io`, `azure.github.io`). These are **reference links embedded in the report output only** — the module never fetches them at runtime, so they are harmless offline (they will simply be non‑navigable from inside the enclave). No action required.

### 2.3 PowerShell Gallery (Class C — MUST neutralize)

| Call | File / Line | Purpose |
|------|-------------|---------|
| `Find-Module` self‑version check + `throw 'Module is out of date.'` | [wara.psm1](wara.psm1#L158), [analyzer/analyzer.psm1](analyzer/analyzer.psm1#L36), [reports/reports.psm1](reports/reports.psm1#L49) | Blocks entry points when the Gallery is unreachable (the `Find-Module` call throws / hangs). |
| `Install-Module -Name ImportExcel` | [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L164), [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1#L163) | Auto‑installs dependency at runtime. |
| `Install-Module -Name powershell-yaml` + `git --version` | [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L172) | **Currently commented out** (dormant). Harmless today, but if these blocks are ever re‑enabled they add a second Gallery install and an external `git` dependency — both would break offline. Flagged so they are not reactivated for USNAT. |

### 2.4 Hard‑coded commercial endpoints / defaults (Class D)

| Item | File / Line | Problem |
|------|-------------|---------|
| `$ResourceURL = "https://management.azure.com/"` default | [advisor/advisor.psm1](advisor/advisor.psm1#L293) | Commercial ARM host as default. Overridden by collector, but wrong for direct calls. |
| `$AzureEnvironment = 'AzureCloud'` default | [wara.psm1](wara.psm1#L145 ), [utils/utils.psm1](utils/utils.psm1#L403) | Defaults to commercial cloud; also drives the `Connect-WAFAzure` reconnect comparison. |
| `ValidatePattern('^https:\/\/.+$')` on data URIs | [wara.psm1](wara.psm1#L145), [wara.psm1](wara.psm1#L151) | Rejects local file paths, forcing an `https` source. |

### 2.5 Non‑endpoint offline considerations

| Item | File / Line | Note |
|------|-------------|------|
| `ImportExcel` module required (community/EPPlus, **not** Microsoft‑authored) | [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L161), [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1#L160) | Must be pre‑staged on the host (no Gallery). If a Microsoft‑only supply chain is required, **prefer the split‑topology approach in §3.7** — it keeps `ImportExcel` entirely out of the enclave rather than replacing it. |
| PowerPoint COM automation | [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1#L1342) | Requires desktop **Microsoft Office** installed on the host. Not an endpoint problem, but note for the reporting stage. |
| `.pptx` template | [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1#L76) | Shipped in the module; already local. ✅ |
| Authenticode signature blocks | every `.psm1` | Editing signed files invalidates the Microsoft signature. In USNAT you will likely need to **re‑sign** with an approved code‑signing cert, or set an appropriate `Execution Policy`. See Risk R4. |

---

## 3. Required changes (with code)

The strategy: introduce **local‑file‑or‑URI resolution** for the data files, make the **version check non‑fatal/offline‑aware**, and **derive the Azure environment from `Get-AzContext`**.

### 3.1 Bundle the data files and add an offline resolver

The files are already in [offline-data/](offline-data/). Add a small helper to `utils/utils.psm1` that accepts either a URL or a local path (defaulting to the bundled copy):

```powershell
function Get-WAFDataSource {
    <#
      .SYNOPSIS
        Returns raw content from a local file (preferred, air-gap safe) or a URI.
      .DESCRIPTION
        If -Path points to an existing file it is read from disk. Otherwise, if it
        looks like an http(s) URI, it is downloaded. This lets USNAT / disconnected
        deployments point at bundled files under .\offline-data without internet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $Source
    )

    if (Test-Path -LiteralPath $Source -PathType Leaf) {
        Write-Debug "Reading data source from local file: $Source"
        return Get-Content -LiteralPath $Source -Raw
    }
    elseif ($Source -match '^https?://') {
        Write-Debug "Downloading data source from URI: $Source"
        return Invoke-RestMethod -Uri $Source
    }
    else {
        throw "Data source [$Source] is neither an existing file nor an http(s) URI."
    }
}
```

Then change the collector parameter defaults and validation in [wara.psm1](wara.psm1#L145) so they point at the bundled files and accept a path **or** a URI:

```powershell
        # was: [ValidatePattern('^https:\/\/.+$')]
        #      [string] $RecommendationDataUri = 'https://azure.github.io/WARA-Build/objects/recommendations.json',
        [ValidateScript({ ($_ -match '^https?://') -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
        [string] $RecommendationDataUri = (Join-Path $PSScriptRoot 'offline-data/recommendations.json'),

        # was: [ValidatePattern('^https:\/\/.+$')]
        #      [string] $RecommendationResourceTypesUri = 'https://azure.github.io/WARA-Build/objects/WARAinScopeResTypes.csv'
        [ValidateScript({ ($_ -match '^https?://') -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
        [string] $RecommendationResourceTypesUri = (Join-Path $PSScriptRoot 'offline-data/WARAinScopeResTypes.csv')
```

Update the two fetch sites in [wara.psm1](wara.psm1#L267) and [wara.psm1](wara.psm1#L280):

```powershell
    # Recommendations (line ~267)
    $RecommendationObject = Get-WAFDataSource -Source $RecommendationDataUri | ConvertFrom-Json

    # Resource types CSV (line ~280)
    $RecommendationResourceTypes = Get-WAFDataSource -Source $RecommendationResourceTypesUri
    $RecommendationResourceTypes = $RecommendationResourceTypes | ConvertFrom-Csv | Where-Object { $_.WARAinScope -eq 'yes' }
```

> Note: `Invoke-RestMethod` on the JSON currently auto‑deserializes; `Get-Content -Raw` does not, so add `| ConvertFrom-Json` as shown. For the CSV, the current code already pipes to `ConvertFrom-Csv`, so no behavior change.

Apply the identical treatment to the analyzer at [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L228) and [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L1165), and to the analyzer defaults at [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L31) / [analyzer/analyzer.psm1](analyzer/analyzer.psm1#L27). The analyzer is a separate script; either dot‑source the helper or inline the same local‑or‑URI logic. Because the analyzer runs standalone, point its default at `$PSScriptRoot/../offline-data/…`.

**Alternative (host internally):** instead of bundling, host the two files on an internal USNAT web server or a storage‑account blob and pass the internal `https://` URL to `-RecommendationDataUri` / `-RecommendationResourceTypesUri`. The resolver above supports both with no further change.

### 3.2 Make the self‑version check offline‑safe

The `Find-Module` call throws when the Gallery is unreachable, which aborts the whole run. Wrap it and add a `-SkipVersionCheck` switch. Applies to [wara.psm1](wara.psm1#L156), [analyzer/analyzer.psm1](analyzer/analyzer.psm1#L34), [reports/reports.psm1](reports/reports.psm1#L47).

```powershell
    param(
        # ... existing params ...
        [switch] $SkipVersionCheck
    )

    if (-not $SkipVersionCheck) {
        try {
            Write-Host 'Checking Version..' -ForegroundColor Cyan
            $LocalVersion   = (Get-Module -Name $MyInvocation.MyCommand.ModuleName).Version
            $GalleryVersion = (Find-Module -Name $MyInvocation.MyCommand.ModuleName -ErrorAction Stop).Version
            if ($LocalVersion -lt $GalleryVersion) {
                Write-Warning "A newer version ($GalleryVersion) is available (installed: $LocalVersion)."
            }
        }
        catch {
            Write-Verbose "Version check skipped (module repository not reachable): $($_.Exception.Message)"
        }
    }
```

Key differences from the current code: (a) `try/catch` so an unreachable Gallery does not throw, and (b) do **not** `throw 'Module is out of date.'` — in a disconnected enclave the operator controls the version manually. A `-SkipVersionCheck` switch lets automation bypass it entirely.

### 3.3 Pre‑stage `ImportExcel` instead of `Install-Module`

At [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1#L162) and [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1#L161):

```powershell
    $ImportExcel = Get-Module -Name ImportExcel -ListAvailable -ErrorAction SilentlyContinue
    if ($null -eq $ImportExcel) {
        # was: Install-Module -Name ImportExcel -Force -SkipPublisherCheck
        throw "The 'ImportExcel' module is not installed. In a disconnected environment, " +
              "pre-stage it (e.g. Save-Module -Name ImportExcel -Path <offline> on a connected host, " +
              "copy to a PSModulePath folder) or install from your internal PSRepository."
    }
```

Pre‑stage on a connected machine and copy it in:

```powershell
# On an internet-connected machine:
Save-Module -Name ImportExcel -Path .\offline-modules
# Copy .\offline-modules\ImportExcel to a folder on $env:PSModulePath in USNAT
```

Or register an internal repository once on the USNAT host and keep `Install-Module` pointed at it:

```powershell
Register-PSRepository -Name Internal -SourceLocation https://<internal-nuget-feed>/nuget -InstallationPolicy Trusted
Install-Module -Name ImportExcel -Repository Internal
```

### 3.4 Derive the Azure environment from `Get-AzContext` (Class D)

This is the change the requirement specifically calls for. Refactor `Connect-WAFAzure` in [utils/utils.psm1](utils/utils.psm1#L401) so that, when the caller does not specify an environment, it **inherits the environment of the existing context** rather than defaulting to `AzureCloud`:

```powershell
function Connect-WAFAzure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [GUID] $TenantID,

        [Parameter(Mandatory = $false)]
        [string] $AzureEnvironment   # no hard-coded default anymore
    )

    $currentContext = Get-AzContext

    # If not specified, inherit the environment already established in the session
    # (USNAT operators will have connected with the correct sovereign environment).
    if ([string]::IsNullOrEmpty($AzureEnvironment)) {
        $AzureEnvironment = $currentContext.Environment.Name ?? 'AzureCloud'
        Write-Debug "AzureEnvironment not supplied; using context environment: $AzureEnvironment"
    }

    if ($currentContext.Tenant.Id -ne $TenantID -or $currentContext.Environment.Name -ne $AzureEnvironment) {
        Write-Debug "Connecting to Azure Tenant $TenantID in environment $AzureEnvironment"
        Connect-AzAccount -Tenant $TenantID -Environment $AzureEnvironment -WarningAction SilentlyContinue | Out-Null
    }
}
```

Correspondingly, in [wara.psm1](wara.psm1#L145) change the collector default so an unspecified value means "use the context":

```powershell
        # was: [string] $AzureEnvironment = 'AzureCloud',
        [string] $AzureEnvironment = ((Get-AzContext).Environment.Name),
```

> Why this matters: the current `Connect-WAFAzure` compares `(Get-AzContext).Environment.Name -ne $AzureEnvironment`. If an operator is already signed into USNAT (e.g. environment name `USNat`) but the default `AzureCloud` is used, the function tries to **reconnect to the commercial cloud** and fails. Inheriting from the context removes that failure and satisfies "get the endpoints from `Get-AzContext`."

Everything downstream already reads the endpoints from the context:
- `$BaseURL = (Get-AzContext).Environment.ResourceManagerUrl` at [wara.psm1](wara.psm1#L300) — ✅ already correct.
- `Get-WAFAdvisorMetadata -ResourceURL $BaseURL` receives the sovereign ARM URL — ✅.

### 3.5 Fix the Advisor‑metadata default host

At [advisor/advisor.psm1](advisor/advisor.psm1#L293), remove the commercial default so a direct call cannot silently hit `management.azure.com`:

```powershell
Function Get-WAFAdvisorMetadata {
    param(
        # was: $ResourceURL = "https://management.azure.com/"
        [string] $ResourceURL = ((Get-AzContext).Environment.ResourceManagerUrl)
    )
    # ...
    # $AdvisorMetadataURI = $ResourceUrl + "providers/Microsoft.Advisor/metadata?api-version=2023-01-01"
```

Guard the string concatenation so a missing trailing slash still produces a valid URI:

```powershell
    $base = $ResourceURL.TrimEnd('/')
    $AdvisorMetadataURI = "$base/providers/Microsoft.Advisor/metadata?api-version=2023-01-01"
```

### 3.6 (Optional) Allow the USNAT environment name explicitly

`Start-WARACollector`'s `-AzureEnvironment` has **no `ValidateSet`**, so a sovereign name such as `USNat` / `USSec` / `AzureUSGovernment` is already accepted. If you later add a `ValidateSet` for UX, be sure to include the USNAT names. Confirm the exact registered name on the target host with:

```powershell
(Get-AzContext).Environment.Name          # e.g. USNat
Get-AzEnvironment | Select-Object Name, ResourceManagerUrl, ActiveDirectoryAuthority
```

Use that exact string for `-AzureEnvironment`, or rely on the inheritance added in §3.4.

### 3.7 Split the collector from the analysis/reporting (removes `ImportExcel` from the enclave) — **recommended**

**Problem.** `ImportExcel` is community‑maintained and wraps the third‑party **EPPlus** library, so it does not satisfy a "Microsoft‑authored modules only" policy. It is also the only reason the enclave needs a non‑Az module pre‑staged. Replacing it in place (Excel COM automation or the Microsoft OpenXML SDK) is a **substantial rewrite** of ~5 analyzer export functions and ~4 report functions, and the hard parts (table styles, data validation, conditional formatting, pivot tables) are exactly what those cmdlets provide for free.

**Key fact that makes this easy.** The stage that actually touches Azure/USNAT endpoints — the **Collector** ([wara.psm1](wara.psm1) `Start-WARACollector`) — has **zero `ImportExcel` dependency**. It only queries Azure and writes a plain **JSON** file. All Excel/PowerPoint work happens later, in stages that never contact Azure:

| Stage | Entry point | `ImportExcel`? | Office COM? | Contacts Azure/USNAT? |
|-------|-------------|----------------|-------------|-----------------------|
| **Collector** | `Start-WARACollector` | ❌ none | ❌ none | ✅ yes — JSON out only |
| **Analyzer** | [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1) | ✅ heavy (EPPlus) | ❌ none | ❌ no |
| **Report** | [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1) | ✅ heavy | ✅ Excel + PowerPoint COM ([line 1389](reports/3_wara_reports_generator.ps1#L1389), [line 1376](reports/3_wara_reports_generator.ps1#L1376)) | ❌ no |

**Approach — run only the Collector inside USNAT; do the Excel/PowerPoint work outside the enclave.**

1. In the disconnected USNAT enclave, run **only** `Start-WARACollector` (with the offline data changes in §3.1). It produces `WARA-File-<timestamp>.json`. No `ImportExcel`, no EPPlus, no Office required inside the enclave.
2. Export that JSON out of the enclave through your approved data‑transfer / review process (it is plain text and easy to inspect for review boards).
3. On a standard, policy‑compliant workstation **outside** the enclave, run `Start-WARAAnalyzer` (→ Excel) and `Start-WARAReport` (→ PowerPoint/Word) against the JSON.

**Why this is the preferred option:**
- **No third‑party module ever enters the enclave** — the policy goal is met by *removal*, not replacement.
- **Zero rewrite** of the Excel/report code — it keeps working as‑is on the outside workstation.
- The JSON boundary is a clean, auditable hand‑off (much easier to review than a binary `.xlsx`).
- It also sidesteps the desktop‑Office requirement (Excel + PowerPoint COM) inside the enclave — see Risk R7.

**Operational notes:**
- The collector's version check and data‑file loading still need the §3.1 / §3.2 changes so it runs cleanly offline.
- If policy forbids moving the JSON out of the enclave, this option does not apply — fall back to replacing `ImportExcel` with Excel COM (Microsoft Office, already a report‑stage dependency) or the Microsoft **OpenXML SDK** (`DocumentFormat.OpenXml`, MIT, no Office needed), both of which are larger efforts.
- Nothing needs to be deleted from the module; you simply do not *invoke* the analyzer/report stages inside USNAT. Optionally document this in the module README so operators know the intended topology.

---

## 4. Bundling / packaging changes

1. Keep the [offline-data/](offline-data/) folder in the module (already created and validated).
2. Declare the bundled files in the manifest so they travel with the package. In [wara.psd1](wara.psd1) uncomment/populate `FileList`:

   ```powershell
   FileList = @(
       'offline-data/recommendations.json',
       'offline-data/WARAinScopeResTypes.csv'
   )
   ```
3. Refresh cadence: these files change as APRL evolves. Establish a manual process — on a connected host run the download, copy the two files into `offline-data/`, and re‑package. (A `Update-WARAOfflineData.ps1` helper that runs the two `Invoke-WebRequest` calls can live outside the disconnected enclave.)

---

## 5. Items that may NOT be fully fixable (risks / open questions)

| ID | Risk | Detail | Mitigation |
|----|------|--------|------------|
| **R1** | **Azure Resource Graph availability** | The entire inventory/recommendation engine depends on `Search-AzGraph` (`Az.ResourceGraph`). If ARG is not deployed/enabled in the USNAT region, `Invoke-WAFQuery` returns nothing and most of the report is empty. | Confirm ARG is available in the target cloud before relying on the module. No code workaround if the service is absent — would require rewriting queries against per‑resource ARM REST calls. |
| **R2** | **Resource Health / Advisor / Support data** | Outages, retirements, service health, support tickets and Advisor recommendations depend on `Microsoft.ResourceHealth`, `Microsoft.Advisor`, `Microsoft.Support` being present and populated in USNAT. Advisor in particular is frequently **not available** in sovereign clouds. | Wrap those collectors in `try/catch` so a `404`/`NotFound` from a missing RP degrades gracefully (empty array) instead of aborting the run. Recommend adding defensive handling in `Get-WAFAdvisorMetadata`, `Get-WAFAdvisorRecommendation`, `Get-WAFOldOutage`, `Get-WAFResourceRetirement`, `Get-WAFSupportTicket`. |
| **R3** | **APRL content relevance** | `recommendations.json` / `WARAinScopeResTypes.csv` are authored for commercial Azure. Some resource types / recommendations may reference services that do not exist in USNAT. | Functionally harmless (extra rows just never match), but the report may list services unavailable in the cloud. Optionally curate a sovereign‑specific subset. |
| **R4** | **Code signing** | Every `.psm1`/`.ps1` carries a Microsoft Authenticode signature. Any edit invalidates it; a locked‑down USNAT host may enforce `AllSigned`. | Re‑sign with an approved internal code‑signing certificate, or set `Set-ExecutionPolicy RemoteSigned`/`Bypass` per policy. This is an operational/compliance decision, not a code fix. |
| **R5** | **Az module version floors** | Manifest requires `Az.Accounts >= 3.0.0`, `Az.ResourceGraph >= 1.0.0` ([wara.psd1](wara.psd1#L53)). Plus `ImportExcel`. These must be pre‑staged in USNAT with versions that actually support the sovereign environment names. | Pre‑stage current Az modules; verify `Get-AzEnvironment` lists the USNAT environment (older Az builds may not know it and would need `Add-AzEnvironment`). `ImportExcel` is only needed by the analyzer/report stages — the **split topology in §3.7** keeps it out of the enclave entirely. |
| **R6** | **Unregistered environment** | If `Get-AzEnvironment` does not include the USNAT cloud, `Connect-AzAccount -Environment` fails and `(Get-AzContext).Environment.Name` inheritance has nothing to inherit. | Operator must register it first via `Add-AzEnvironment` (ARM URL, AAD authority, Graph, etc.) or a provided sovereign discovery endpoint. Document this as a prerequisite. |
| **R7** | **PowerPoint/Word COM for reports** | `Start-WARAReport` drives PowerPoint via COM and needs desktop Office on the host. | Air‑gapped reporting hosts must have Office installed; otherwise run only the Collector in USNAT and generate the Excel/deck elsewhere (the **recommended split topology in §3.7**). Not an endpoint issue. |

---

## 6. Change checklist (by file)

> **Status:** implemented in the commit documented by [USNAT-Refactoring-Changes.md](USNAT-Refactoring-Changes.md). R2 was implemented at the **collector call sites** in [wara.psm1](wara.psm1) rather than inside each leaf module (cleaner central handling + `collectionErrors` output).

- [x] **Deployment topology (§3.7):** confirmed — run **only the Collector** in USNAT; Excel/report stages run outside. Collector path is now free of `ImportExcel`/Gallery.
- [x] [utils/utils.psm1](utils/utils.psm1) — added `Get-WAFDataSource`; refactored `Connect-WAFAzure` to inherit environment from `Get-AzContext` **and prompt to register unknown sovereign environments** (§3.1, §3.4, R6).
- [x] [wara.psm1](wara.psm1) — repointed data URIs to `offline-data`, relaxed validation, used `Get-WAFDataSource`, non‑fatal version check + `-SkipVersionCheck`, `-AzureEnvironment` inherits context, **graceful degradation + `collectionErrors`** (§3.1, §3.2, §3.4, R2).
- [x] [advisor/advisor.psm1](advisor/advisor.psm1) — removed commercial default host (now context), `TrimEnd` slash (§3.5). Missing‑Advisor handling is done at the collector call site (R2).
- [x] [analyzer/analyzer.psm1](analyzer/analyzer.psm1) — non‑fatal version check + `-SkipVersionCheck` (§3.2).
- [x] [analyzer/2_wara_data_analyzer.ps1](analyzer/2_wara_data_analyzer.ps1) — offline data resolution; replaced `Install-Module ImportExcel` with pre‑stage `throw` (§3.1, §3.3).
- [x] [reports/reports.psm1](reports/reports.psm1) — non‑fatal version check + `-SkipVersionCheck` (§3.2).
- [x] [reports/3_wara_reports_generator.ps1](reports/3_wara_reports_generator.ps1) — replaced `Install-Module ImportExcel` with pre‑stage `throw` (§3.3).
- [x] R2 handling — implemented centrally in [wara.psm1](wara.psm1) for Advisor/Outages/Retirements/Support/ServiceHealth (leaf modules left unchanged; failures recorded in `collectionErrors`).
- [x] [wara.psd1](wara.psd1) — populated `FileList`; exported `Get-WAFDataSource`.
- [ ] Re‑sign all edited files or adjust execution policy (R4 — **you are handling this**).

---

## 7. Prerequisite runbook for the USNAT host (operator)

```powershell
# 1. Ensure the sovereign environment is known to Az
Get-AzEnvironment | Where-Object Name -match 'Nat|Sec|Gov'
# If missing, register it (values supplied by your USNAT onboarding):
# Add-AzEnvironment -Name USNat -ARMEndpoint https://management.azure.eaglex.ic.gov/ ...

# 2. Sign in to the sovereign cloud (endpoints now come from the context)
Connect-AzAccount -Environment USNat -Tenant <tenant-guid>
(Get-AzContext).Environment.Name              # verify: USNat
(Get-AzContext).Environment.ResourceManagerUrl

# 3. Pre-stage required modules (copied from a connected host)
#    Az.Accounts (>=3.0.0), Az.ResourceGraph (>=1.0.0), ImportExcel

# 4. Run the collector against bundled offline data (no internet)
Start-WARACollector -TenantID <tenant-guid> -SubscriptionIds <sub> -SkipVersionCheck
#   -RecommendationDataUri / -RecommendationResourceTypesUri default to .\offline-data\*
```

---

## 8. Bottom line

- **Fully fixable with code changes:** all public‑content downloads (bundled), PowerShell Gallery calls (made optional / internal), and every commercial‑cloud default (now derived from `Get-AzContext`). The authenticated Azure surface already routes through the context and needs only defensive hardening.
- **Not a code problem — environmental prerequisites:** ARG/Advisor/ResourceHealth service availability in USNAT (R1/R2), Az environment registration (R6), module pre‑staging (R5), code‑signing policy (R4), and Office for reporting (R7). These must be validated with the USNAT platform team; where a backing service is simply absent, the corresponding section of the report will be empty and that cannot be worked around in code.
```