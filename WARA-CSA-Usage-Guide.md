# WARA Assessment — CSA Run Guide (v1.0.7)

A practical, end-to-end guide for a Cloud Solution Architect (CSA) running the **Well-Architected
Reliability Assessment (WARA)** against a customer environment — including **sovereign / disconnected
clouds (USNAT, USSec, USGov)**.

This guide targets the **`1.0.7`** module in this folder, whose collector talks to Azure over **REST**
and needs **only `Az.Accounts`** (no `Az.ResourceGraph`).

---

## 1. What WARA does (three phases)

| Phase | Command | Input | Output | Needs |
|---|---|---|---|---|
| 1. **Collect** | `Start-WARACollector` | live Azure (as you) | `WARA-File-<timestamp>.json` | `Az.Accounts`, network to Azure ARM |
| 2. **Analyze** | `Start-WARAAnalyzer` *or* `Export-WARAActionPlanExcelCom.ps1` | the JSON | `Expert-Analysis-v1-<timestamp>.xlsx` | Excel workbook engine (see below) |
| 3. **Report** | `Start-WARAReport` | the **reviewed** Excel | PowerPoint + Findings xlsx + CSV | desktop Excel & PowerPoint (Microsoft-only, no ImportExcel) |

Between phases 2 and 3 there is a **mandatory manual review step** (Section 6).

---

## 2. Prerequisites

**All phases**
- **PowerShell 7.4+** (`$PSVersionTable.PSVersion`). Windows only for phases 2–3 (Office COM).
- The WARA module folder (`…\wara\1.0.7`).

**Phase 1 – Collect**
- **`Az.Accounts`** (>= 3.0.0). *`Az.ResourceGraph` is **not** required.*
- You must be able to sign in as a user with at least **Reader** on the target subscriptions.

**Phase 2 – Analyze** (choose ONE workbook engine)
- **Option A – `Start-WARAAnalyzer`**: requires the **`ImportExcel`** module (no desktop Excel needed).
- **Option B – `Export-WARAActionPlanExcelCom.ps1`**: requires **desktop Microsoft Excel** (uses local
  Excel via COM). Microsoft-only — no `ImportExcel`. Use this where third-party modules aren't allowed.

**Phase 3 – Report**
- Desktop **Microsoft Excel + PowerPoint** installed (the generator drives both via COM).
  Microsoft-only — **no `ImportExcel`** required.
- **Clipboard History must be OFF** (Settings ▸ System ▸ Clipboard) — the generator refuses to run with
  it enabled.

**Disconnected / sovereign notes**
- Recommendation data is **bundled offline** (`1.0.7/offline-data/`) — no internet needed for collection.
- Version checks are **best-effort** and are skipped automatically when the PowerShell Gallery is
  unreachable; you can also pass **`-SkipVersionCheck`** everywhere.
- `ImportExcel` is only needed for **Phase 2 Option A** (`Start-WARAAnalyzer`). If you use that engine on
  a disconnected host, **pre-stage** it: on a connected machine run
  `Save-Module -Name ImportExcel -Path <folder>`, copy the folder onto a `PSModulePath` directory of the
  enclave host. Phases 1 and 3 never need it.

---

## 3. One-time setup

### 3a. Import the module
Import explicitly by path (deterministic when multiple copies exist):

```powershell
Import-Module 'C:\Users\<you>\...\PowerShell\Modules\wara\1.0.7\wara.psd1' -Force
```

(If the module is installed normally on `PSModulePath`: `Import-Module wara -RequiredVersion 1.0.7`.)

### 3b. Sign in to Azure
**Commercial:**
```powershell
Connect-AzAccount -Tenant <tenantId>
```

**Sovereign (e.g. USNAT)** — register the environment once, then connect to it:
```powershell
# Register the sovereign cloud (endpoints per your enclave). You already know how to do this.
Add-AzEnvironment -Name 'USNat' -ARMEndpoint 'https://management.azure.<sovereign-host>/'   # example
Connect-AzAccount -Environment 'USNat' -Tenant <tenantId>
```

The collector **inherits the cloud from your `Get-AzContext`**, so once you're connected to the right
environment you don't have to pass endpoints anywhere. If the environment isn't registered, the collector
stops early with guidance to run `Add-AzEnvironment`.

---

## 4. Phase 1 — Collect

Run from the folder where you want the JSON written (it lands in the current directory).

```powershell
Start-WARACollector `
    -TenantID       <tenantId> `
    -SubscriptionIds <subId1>,<subId2> `
    -SkipVersionCheck
```

Common options:
- `-ResourceGroups <rgResourceId…>` — scope to specific resource groups instead of whole subscriptions.
- `-AzureEnvironment 'USNat'` — usually unnecessary (inherited from context); pass it to be explicit.
- Specialized workloads: `-SAP`, `-AVD`, `-AVS`, `-HPC`, `-AI_GPT_RAG`, `-ORACLE`.
- `-PassThru` — return the object instead of writing a file.

**Output:** `.\WARA-File-<yyyy-MM-dd-HH-mm>.json`.

**Graceful degradation:** optional data sources (Advisor, Resource Health outages/retirements, Support,
Service Health) that aren't reachable in the target cloud are **recorded in the JSON's `collectionErrors`
array** instead of failing the whole run. Check that array to see what (if anything) was skipped.

---

## 5. Phase 2 — Analyze (build the Expert-Analysis workbook)

Pick the engine that matches your environment:

**Option A — ImportExcel (no desktop Excel):**
```powershell
Start-WARAAnalyzer -JSONFile '.\WARA-File-2026-07-16-20-01.json' -SkipVersionCheck
```

**Option B — local Excel COM (Microsoft-only, no ImportExcel):**
```powershell
& 'C:\...\wara\1.0.7\analyzer\Export-WARAActionPlanExcelCom.ps1' `
    -JSONFile '.\WARA-File-2026-07-16-20-01.json'
```

**Output (either):** `Expert-Analysis-v1-<timestamp>.xlsx` — the working "action plan" workbook with the
`4.ImpactedResourcesAnalysis`, `2.WorkloadInventory`, `3.AnalysisPlanning`, `5.PlatformIssuesAnalysis`,
and `6.SupportRequestsAnalysis` sheets populated.

---

## 6. Phase 3 (manual) — Review the findings  ⚠️ required

The report generator **will not produce output** until the High/Medium findings are reviewed.

1. Open the `Expert-Analysis-v1-*.xlsx` workbook in Excel.
2. Go to sheet **`4.ImpactedResourcesAnalysis`**.
3. In column **A (`REQUIRED ACTIONS / REVIEW STATUS`)**, change every **non-Low** row from its
   `REQUIRED ACTIONS: …` prompt to **`Reviewed`** (there's a dropdown). Low-impact rows are exempt.
4. **Save**, then **close** the workbook (leave it open and the next step fails — Excel locks the file).

---

## 7. Phase 4 — Report (PowerPoint + Findings + CSV)

```powershell
Start-WARAReport `
    -ExpertAnalysisFile '.\Expert-Analysis-v1-2026-07-16-20-01.xlsx' `
    -CustomerName 'Contoso' `
    -WorkloadName 'Contoso Prod Platform' `
    -PPTTemplateFile 'C:\...\wara\1.0.7\reports\Mandatory - Executive Summary presentation - Template.pptx' `
    -includeLow `
    -SkipVersionCheck
```

- **Do NOT pass `-AssessmentFindingsFile`.** Leave it out so the generator uses the shipped template
  (`1.0.7\reports\Assessment-Findings-Report-v1.xlsx`). Passing a *previously generated* findings file
  causes "pivot/chart already exists" errors.
- `-includeLow` includes Low-impact items in the report body.

**Outputs (in the current folder):**
- `Executive Summary Presentation - <Customer> - <timestamp>.pptx`
- `Assessment-Findings-Report-v1-<timestamp>.xlsx`
- `Impacted Resources and Recommendations Template <timestamp>.csv`

---

## 8. Troubleshooting (most common)

| Symptom | Cause | Fix |
|---|---|---|
| `Start-WARAReport` finishes with **no output** | The review gate: non-Low findings still say `REQUIRED ACTIONS: …` | Mark all non-Low rows `Reviewed` (Section 6), save, close, re-run |
| **"file … is being used by another process"** | The Expert-Analysis workbook is still open in Excel | Save & close it, then re-run |
| **"Name already exists in the drawings collection"** / pivot "already exists" | You passed a **prior output** as `-AssessmentFindingsFile` | Omit `-AssessmentFindingsFile` (uses the shipped template) |
| Report refuses to start, mentions **Clipboard History** | Windows Clipboard History is enabled | Turn it off (Settings ▸ System ▸ Clipboard) |
| Collector error: environment **not registered** | Sovereign cloud not added | `Add-AzEnvironment …` then `Connect-AzAccount -Environment …` |
| A data section is empty & listed in **`collectionErrors`** | That RP/API isn't available in the cloud | Expected in some sovereign clouds; the run still completes |
| `ImportExcel` "not installed" during Analyze | Module missing on a disconnected host (Phase 2 Option A only) | Pre-stage it (`Save-Module ImportExcel`), or use the Excel-COM analyzer for Phase 2. Phase 3 no longer needs it |

Leftover hidden `EXCEL.EXE` / `POWERPNT.EXE` automation processes are cleaned up automatically (only the
`/automation` instances the tools spawn — your own open Office windows are never touched).

---

## 9. End-to-end example (copy/paste)

```powershell
# 0) Setup
Import-Module 'C:\...\wara\1.0.7\wara.psd1' -Force
Connect-AzAccount -Tenant <tenantId>            # add -Environment 'USNat' for sovereign
Set-Location 'C:\WARA\Contoso'                  # outputs land here

# 1) Collect
Start-WARACollector -TenantID <tenantId> -SubscriptionIds <subId> -SkipVersionCheck

# 2) Analyze (choose one)
Start-WARAAnalyzer -JSONFile .\WARA-File-*.json -SkipVersionCheck
# or, Microsoft-only:
# & 'C:\...\wara\1.0.7\analyzer\Export-WARAActionPlanExcelCom.ps1' -JSONFile .\WARA-File-*.json

# 3) REVIEW in Excel: set column A to 'Reviewed' for non-Low rows on '4.ImpactedResourcesAnalysis'; save & close.

# 4) Report
Start-WARAReport -ExpertAnalysisFile .\Expert-Analysis-v1-*.xlsx `
    -CustomerName 'Contoso' -WorkloadName 'Contoso Prod' `
    -PPTTemplateFile 'C:\...\wara\1.0.7\reports\Mandatory - Executive Summary presentation - Template.pptx' `
    -includeLow -SkipVersionCheck
```

---

## 10. What's different in this build (1.0.7)

- **Collector is REST-based** and requires **only `Az.Accounts`** (auth/endpoints/token) — no
  `Az.ResourceGraph`. Works across sovereign clouds by inheriting the cloud from `Get-AzContext`.
- **Offline-capable**: recommendation data bundled under `1.0.7/offline-data/`; version checks are
  best-effort / skippable.
- **Optional data sources fail gracefully** into `collectionErrors` rather than aborting collection.
- **Two Phase-2 engines**: `Start-WARAAnalyzer` (ImportExcel) or `Export-WARAActionPlanExcelCom.ps1`
  (Microsoft-only, local Excel) — both produce the same workbook.
- **Phase-3 report is Microsoft-only**: `Start-WARAReport` now drives desktop **Excel + PowerPoint** via
  COM for the Findings workbook (tables, pivots, charts) — **no `ImportExcel`** dependency.

*See the design notes in `1.0.7/REST-Collector-Compromise-AzAccounts.md`,
`1.0.7/USNAT-Refactoring-Changes.md`, and the cross-cloud backlog in
`1.0.7/REST-Collector-CrossCloud-TODO.md`.*
