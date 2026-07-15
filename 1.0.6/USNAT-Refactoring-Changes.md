# WARA Module — USNAT / Sovereign-Cloud Refactoring: Change Log

Companion to [USNAT-Refactoring-Assessment.md](USNAT-Refactoring-Assessment.md). This file records **what was actually changed and why**, plus anything discovered during implementation. It is kept in sync with the code — if a change is later corrected, this log is updated too.

**Design goal:** make the **Collector** run in *any* registered sovereign Azure cloud (starting with USNAT) with **no public internet** and **no PowerShell Gallery**, taking all Azure endpoints from `Get-AzContext`. Endpoints are never hard-coded, so the same code works for USNAT, USSec, USGov, China, or commercial.

## Decisions applied (from the R1–R7 discussion)

| Ref | Decision | How it was implemented |
|-----|----------|------------------------|
| **R1** | Keep using `Search-AzGraph`; do **not** rewrite ARG to REST for now | Left `Invoke-WAFQuery` / `Search-AzGraph` untouched. Token/endpoint acquisition continues to flow through the Az context. |
| **R2** | Try every optional RP; if unreachable, note it in the output instead of failing collection | Wrapped Advisor, Outages, Retirements, Support, Service Health in `try/catch`; failures are recorded in a new `collectionErrors` array in the output JSON and the run continues. |
| **R3** | Accept APRL-content mismatch | No code change (documented as accepted). |
| **R4** | Code signing handled by you | Files were edited (signatures now invalid); **not** re-signed. No action taken per your instruction. |
| **R5** | Use JSON instead of Excel | The Collector already emits **JSON only** and needs no `ImportExcel`. Removed all Gallery auto-install so the Collector path never requires a non-Microsoft module. Excel/report stages run outside the enclave (Option A, §3.7 of the assessment). Impairment analysis below. |
| **R6** | Prompt to `Add-AzEnvironment` if needed | `Connect-WAFAzure` now validates the target environment against `Get-AzEnvironment` and stops with actionable registration guidance if it is unknown. |
| **R7** | PowerPoint/Word will be installed | No change needed; report stage keeps using Office COM. |

---

## Changes by file

### 1. `utils/utils.psm1`

**a) New function `Get-WAFDataSource`.**
*What:* Returns the content of a data source that is **either a local file path or an http(s) URI** — local file is read with `Get-Content -Raw`, otherwise the URI is downloaded with `Invoke-RestMethod`.
*Why:* Lets the Collector load the bundled APRL data (`offline-data/`) with zero network access, while still allowing an internal URL to be passed. Central, testable resolver used by the Collector.

**b) Refactored `Connect-WAFAzure`.**
*What:*
- Removed the hard-coded `$AzureEnvironment = 'AzureCloud'` default.
- If the caller passes no environment, it **inherits `(Get-AzContext).Environment.Name`** (falls back to `AzureCloud` only when there is no context at all).
- Validates the requested environment against `@((Get-AzEnvironment).Name)` and, if unknown, **throws with `Add-AzEnvironment` guidance** (R6).
- Connects only when the tenant/environment differs from the current context.

*Why:* This is the core multi-sovereign change. The operator signs in to the desired cloud once; the module follows that context instead of forcing commercial endpoints. The registration check turns a confusing downstream endpoint failure into a clear, fixable message.

### 2. `utils/utils.psd1`

*What:* Added `Get-WAFDataSource` to `FunctionsToExport`. Removed `Get-WAFAzureEnvironmentAPIUri` from the export list.
*Why:* Export the new resolver so the analyzer/session can use it. `Get-WAFAzureEnvironmentAPIUri` was listed but **had no backing function anywhere in the module** (a latent dead entry) — removing it keeps the manifest honest.

### 3. `wara.psm1` (`Start-WARACollector`)

**a) Parameters.**
- `-AzureEnvironment` default changed `'AzureCloud'` → `''` (empty = inherit from context, per R6/multi-sovereign).
- `-RecommendationDataUri` / `-RecommendationResourceTypesUri`: `ValidatePattern('^https://…')` → `ValidateScript` that accepts an http(s) URI **or** an existing local file; defaults repointed to the bundled `offline-data/` copies via `Join-Path $PSScriptRoot`.
- Added `-SkipVersionCheck` switch.

**b) Version check.**
*What:* The old block called `Find-Module` and `throw 'Module is out of date.'`. Now wrapped in `try/catch` (Gallery unreachable → `Write-Verbose`, no throw), downgraded the "newer version" message to `Write-Warning`, and gated the whole thing behind `-SkipVersionCheck`.
*Why:* `Find-Module` hangs/throws with no internet, which previously aborted the entire collection in an air-gapped enclave.

**c) Data loading.**
*What:* `Invoke-RestMethod $RecommendationDataUri` → `Get-WAFDataSource` (+ `ConvertFrom-Json` when a local string is returned). Same for the resource-types CSV.
*Why:* Air-gap-safe loading of the bundled data.

**d) Environment/endpoint resolution after connect.**
*What:* After `Connect-WAFAzure`, the code now sets `$AzureEnvironment = (Get-AzContext).Environment.Name` and `$BaseURL = (Get-AzContext).Environment.ResourceManagerUrl`.
*Why:* Guarantees every downstream call (Advisor metadata REST, etc.) uses the **actual sovereign endpoints**, and the recorded `scriptDetails.AzureEnvironment` reflects the real cloud even when the parameter was left empty.

**e) Graceful degradation (R2).**
*What:* Introduced `$collectionErrors = [System.Collections.Generic.List[object]]::new()`. Wrapped these optional collectors in `try/catch`, each defaulting to `@()` on failure and appending `{ Source; Reachable=$false; Error }` to `$collectionErrors`:
- Advisor **metadata** (`Get-WAFAdvisorMetadata`)
- Advisor **recommendations** (`Get-WAFAdvisorRecommendation`)
- **Outages** (`Get-WAFOldOutage`)
- **Retirements** (`Get-WAFResourceRetirement`)
- **Support tickets** (`Get-WAFSupportTicket`)
- **Service health** (`Get-WAFServiceHealth`)

Also made `Get-WARAOtherRecommendations` run only when Advisor metadata was successfully retrieved (otherwise `@()`), since it depends on the metadata.

*Why:* Sovereign clouds may not expose every resource provider (Advisor especially). Per R2, a missing RP now degrades to an empty section plus an audit record, instead of aborting the whole collection.

**f) Output JSON.**
*What:* Added `collectionErrors = $collectionErrors` as a new top-level property of the output object.
*Why:* Surfaces exactly which optional data sources could not be reached, for review and for downstream tooling. Existing consumers are unaffected (additive property).

> **Note on core queries:** `Invoke-WAFQuery` / `Invoke-WAFQueryLoop` (the main ARG inventory + APRL recommendation queries) were **intentionally not** wrapped. Per R1, Resource Graph is available in USNAT; if it truly fails, collection cannot meaningfully proceed, so a hard failure there is correct.

### 4. `advisor/advisor.psm1` (`Get-WAFAdvisorMetadata`)

*What:*
- Default `$ResourceURL` changed from `"https://management.azure.com/"` to `(Get-AzContext).Environment.ResourceManagerUrl`.
- URI build hardened: `$ResourceUrl.TrimEnd('/') + "/providers/Microsoft.Advisor/metadata?api-version=2023-01-01"`.

*Why:* Removes the only hard-coded commercial host so a direct call cannot silently hit commercial ARM; `TrimEnd` tolerates sovereign ARM URLs with/without a trailing slash.

### 5. `analyzer/analyzer.psm1` & `reports/reports.psm1` (wrappers)

*What:* Same version-check treatment as the Collector: `try/catch`, non-fatal, `Write-Warning` instead of `throw`, plus a `-SkipVersionCheck` switch. The switch is **removed from `$PSBoundParameters` before splatting** to the child `.ps1` (which does not define it). Relaxed the analyzer wrapper's `-RecommendationDataUri` validation to allow a local path.
*Why:* Same disconnected-Gallery resilience for the analyze/report entry points.

### 6. `analyzer/2_wara_data_analyzer.ps1`

*What:*
- Default `-RecommendationDataUri` and `$RecommendationResourceTypesUri` repointed to `../offline-data/…` with a local-or-URI validation/loader (mirrors the Collector).
- `Get-WARARecommendationList` and the resource-types CSV import now read the local file when present, else download.
- **Replaced `Install-Module -Name ImportExcel …` with a `throw`** instructing the operator to pre-stage the module (Save-Module + copy, or internal PSRepository).

*Why:* No Gallery auto-install; supports bundled data. (`ImportExcel` itself is still used here — this stage is intended to run **outside** the enclave per Option A; the throw simply prevents a silent Gallery call and documents the Microsoft-only pre-staging path.)

### 7. `reports/3_wara_reports_generator.ps1`

*What:* Replaced `Install-Module -Name ImportExcel …` with the same pre-stage `throw`.
*Why:* Same as above. The report stage additionally uses Excel/PowerPoint COM (Microsoft Office), which you confirmed will be available (R7).

### 8. `wara.psd1`

*What:* Added `Get-WAFDataSource` to `FunctionsToExport`; populated `FileList` with the two `offline-data/` files.
*Why:* Expose the resolver; ensure the bundled data travels with the package.

---

## New output shape: `collectionErrors`

`Start-WARACollector` output JSON gains a top-level array. When everything succeeds it is empty (`[]`). Example when Advisor is absent in a sovereign cloud:

```json
"collectionErrors": [
  { "Source": "AdvisorMetadata",        "Reachable": false, "Error": "Response status code does not indicate success: 404 (NotFound)." },
  { "Source": "AdvisorRecommendations", "Reachable": false, "Error": "..." }
]
```

`Source` values: `AdvisorMetadata`, `AdvisorRecommendations`, `Outages`, `Retirements`, `SupportTickets`, `ServiceHealth`.

---

## R5 impairment findings (JSON instead of Excel)

You asked to be told if relying on JSON rather than the Excel workbook impairs anything. Findings:

1. **Collection is unaffected.** The Collector already produces a complete JSON document (`impactedResources`, `resourceType`, `advisory`, `outages`, `retirements`, `supportTickets`, `serviceHealth`, `resourceInventory`, and now `collectionErrors`). No data is lost by treating the JSON as the enclave deliverable.
2. **What the Excel stage adds (and you forgo if you never run it):** the analyzer's `.xlsx` is an **interactive expert-review artifact** — dropdown data validation (`Pending/Reviewed`, `High/Medium/Low`), colour-coded conditional formatting, tables, and pivot tables — designed for a consultant to fill in *between* analyze and report. JSON has none of that review UX.
3. **Recommended workflow (Option A):** run only the Collector in USNAT → export the JSON through your review process → run `Start-WARAAnalyzer` / `Start-WARAReport` on a workstation **outside** the enclave (with Office + a pre-staged `ImportExcel`). This keeps the third-party module out of the enclave entirely and preserves the full Excel/PowerPoint deliverables.
4. **If you require the Excel to be produced *inside* the enclave with Microsoft-only modules**, that is a larger effort (rewrite the analyzer/report Excel layer onto Excel COM or the Microsoft OpenXML SDK) — not done here. See assessment §3.7.

**Net:** no impairment to *data collection*; the only thing gated behind Excel is the human-in-the-loop review workbook, which Option A relocates rather than loses.

---

## Validation performed

- **Syntax:** all 9 edited files parse cleanly (`[System.Management.Automation.Language.Parser]::ParseFile`).
- **Manifest:** `Test-ModuleManifest wara.psd1` succeeds (17 exported functions).
- **Bundled data:** `offline-data/recommendations.json` (393 records) and `offline-data/WARAinScopeResTypes.csv` (3,861 rows) exist and parse.
- **Not run:** a live `Start-WARACollector` against a real sovereign tenant (requires USNAT access). Recommend a smoke test there.

## Not done (by design / your instruction)

- **Re-signing** the edited files (R4 — you will handle).
- **Rewriting ARG queries to REST** (R1 — deferred).
- **Replacing `ImportExcel` in-enclave** (R5 — Option A relocation preferred; in-enclave Microsoft-only Excel would be a separate rewrite).
- **`ConfigFile` AzureEnvironment:** the config-file path still does not read an `AzureEnvironment` key; with the new inherit-from-context behavior it is no longer required (leave unset to follow the signed-in cloud). Flagged here in case you want an explicit key later.
