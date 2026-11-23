# Group SOA Conversion Tool - v1.08

> This folder contains **v1.08**, the enterprise version of the tool (large tenants, batching, CSV waves). For the original, simpler version see [v1.03](../v1.03/README.md). The [root README](../README.md) compares the versions.

A PowerShell GUI tool for converting the *Group Source of Authority (SOA)* between on-premises Active Directory and Microsoft Entra ID (cloud) by toggling `isCloudManaged` via the Microsoft Graph `onPremisesSyncBehavior` API.

v1.08 lists **on-premises synced groups** and **groups whose SOA is already converted** to the cloud:
- Security Groups (not mail-enabled)
- Mail-Enabled Security Groups
- Distribution Groups

Cloud-only groups and Microsoft 365 (Unified) groups are never listed.

This is intended to support the approach described in:
- [Guidance for Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/concept-group-source-of-authority-guidance)
- [Configure Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure)

## Features

- **Built for large tenants (15K+ synced groups)**:
  - Server-side Graph filters.
  - JSON batching (20 requests per call) for SOA status, nesting and convert / rollback.
  - Retry-After aware throttling handling.
- **Synced and converted groups**: discovery runs two queries and merges them, so converted groups stay visible and can be rolled back.
- **Search scope tick boxes**: Security / Mail-enabled security / Distribution lists. None ticked = all types.
- **Optional on-demand nesting analysis**: skip it on Search, and it runs only for the groups you select for an operation.
- **Filters**: group type, free-text search (display name / email, applied 400 ms after typing or on Enter) and CSV import.
- **CSV-driven waves**:
  - Load a list of groups straight from Graph by ObjectId, Mail or DisplayName.
  - Or filter the list that is already loaded.
- **Select all**: operate on every group in the filtered list across all pages.
  - Only groups in the right SOA state are processed.
  - The box clears itself when you click a row.
- **Batches per nesting depth level**:
  - Conversion runs bottom-up (children first); rollback runs top-down (parents first).
  - Each level is completed before the next one starts.
- **Honest results**:
  - Failed sub-requests are verified with a GET.
  - Each group is reported as **Success**, **Failed**, **Unverified** (state could not be confirmed) or **Skipped** (cancelled).
- **Per-run CSV report** in the `Reports` folder.
- **Progress with ETA** and a **Cancel job** button. The window can't be closed while a job is running.
- **Nested group warnings**: warns if a parent is converted before its unconverted children.
- **CSV export** of the current list, **log file** in the `Logs` folder, and an **Open Log File** button.
- **Automatic module installation and permission consent flow**, as in v1.03.

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Graph PowerShell module: `Microsoft.Graph.Groups` (installed automatically in *CurrentUser* scope if missing)
- Microsoft Entra Connect Sync **2.5.76.0** or later, or Microsoft Entra Cloud Sync **1.1.1370.0** or later
- Roles and permissions:
  - **Hybrid Administrator** role.
  - `Group.ReadWrite.All` and `Group-OnPremisesSyncBehavior.ReadWrite.All` consented for *Microsoft Graph Command Line Tools*.

  See the [root README](../README.md#permissions--roles) for how to grant consent.

## Usage

1. **Run the script**:
   ```powershell
   .\Group-SOA-Conversion-Tool_1.08.ps1
   .\Group-SOA-Conversion-Tool_1.08.ps1 -TenantId "00000000-0000-0000-0000-000000000000"
   ```
2. **Connect to Graph**. Sign in, and consent is requested if needed. Groups aren't loaded automatically.
3. **Choose the search scope (optional)**:
   - Tick one or more group types; none ticked = all types.
   - Optionally tick **Skip nesting analysis on Search** for very large tenants.
4. **Search Groups**. The tool runs both discovery queries, reads the SOA status in batches and (unless skipped) analyzes nesting.
5. **Filter (optional)**:
   - **Group type**: narrow the list to one type.
   - **Search**: free text on display name / email.
   - **Import CSV...**: supported columns are `ObjectId`, `Id`, `GroupId`, `Mail`, `Email`, `EmailAddress`, `PrimarySmtpAddress`, `DisplayName`, `Name`, `GroupName`, `Value`, or a plain single-column list. Both `,` and `;` delimiters work.
     - If no groups are loaded yet, the groups in the CSV are loaded directly from Graph.
     - If groups are already loaded, you choose: **Yes** = load the CSV groups fresh from Graph (replaces the list), **No** = filter the current list, **Cancel** = do nothing.
     - CSV discovery applies the search scope tick boxes. Entries that are not found, ambiguous, cloud-only / Microsoft 365, or out of scope are listed in a summary and in the log.
   - **Clear Filters** resets the type, search and CSV filters.
6. **Select groups**:
   - Select rows (multi-select).
   - Or tick **Select all groups in the list (all pages)**. Convert then uses every filtered group with Cloud Managed = False, and Rollback every filtered group with Cloud Managed = True. Clicking a row clears the tick box.
7. **Convert to Cloud Managed** or **Roll Back to On-Prem**:
   - Nesting is analyzed on demand if needed, and nested-order warnings are shown.
   - Confirm the list, then follow the progress and ETA.
   - **Cancel job** stops after the current batch.
   - A summary and a report file are shown at the end.
   - Before rolling back, remove cloud users from the groups and remove the groups from access packages. The rollback is only complete after the next sync cycle.
8. **Export List to CSV**, **Open Log File**, **Disconnect from Graph** (clears all loaded groups) or **Exit**.

## Understanding Nested Groups

- If **GroupB** is a member of **GroupA**, GroupB is a child and GroupA is a parent. Convert GroupB **before** GroupA.
- **Nesting Depth**:
  - 0 = no nested children in scope.
  - 1 = has direct nested children.
  - 2+ = deeper nesting.
- **Conversion** runs from depth 0 upwards. **Rollback** runs from the highest depth downwards.

```
GroupA (Depth 2)
  └── GroupB (Depth 1)  ← member of GroupA
        └── GroupC (Depth 0)  ← member of GroupB
```

**Conversion order**: GroupC → GroupB → GroupA. **Rollback order**: GroupA → GroupB → GroupC.

## Output Files

| Location | File | Content |
| --- | --- | --- |
| `Logs\` | `GroupSOAConversion_YYYYMMDD_HHMM.log` | Session log (falls back to the script folder if `Logs` can't be created) |
| `Reports\` | `GroupSOA_Convert_YYYYMMDD_HHMMSS.csv` / `GroupSOA_Rollback_YYYYMMDD_HHMMSS.csv` | One row per group processed in a run |
| script folder | `GroupSOAExport_YYYYMMDD_HHMMSS.csv` | Export of the current list (all pages, filters and sort applied) |

**Report columns**:
- `Timestamp`
- `Operation` (`Convert` / `Rollback`)
- `DisplayName`
- `Mail`
- `GroupType`
- `ObjectId`
- `Depth`
- `Result` (`Success`, `Failed`, `Unverified`, `Skipped`)
- `Detail`

**Export columns**: `DisplayName`, `Mail`, `GroupType`, `IsCloudManaged` (`True` / `False` / `Unknown`), `NestingDepth`, `SecurityEnabled`, `MailEnabled`, `ObjectId`.

These files contain tenant data. The repository's `.gitignore` excludes them.

## Graph API Calls

**Discovery**, query 1: on-premises synced groups that are not converted yet. It's sent with the header `ConsistencyLevel: eventual`:
```
GET https://graph.microsoft.com/v1.0/groups?$filter=onPremisesSyncEnabled eq true&$count=true&$select=...&$top=999
```

**Discovery**, query 2: groups whose SOA is already converted. Same header:
```
GET https://graph.microsoft.com/v1.0/groups?$filter=onPremisesSyncBehavior/isCloudManaged eq true&$count=true&$select=...&$top=999
```

**Batching**. SOA status, nesting (`/groups/{id}/members/microsoft.graph.group`), CSV lookups and convert / rollback are sent as up to 20 sub-requests per call:
```
POST https://graph.microsoft.com/v1.0/$batch
```

**Sub-requests**:
```
GET   /groups/{ID}/onPremisesSyncBehavior?$select=isCloudManaged
PATCH /groups/{ID}/onPremisesSyncBehavior   { "isCloudManaged": true }    (convert)
PATCH /groups/{ID}/onPremisesSyncBehavior   { "isCloudManaged": false }   (rollback)
```

## Why two discovery queries?

According to Microsoft's [status of attributes after you convert SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure#status-of-attributes-after-you-convert-soa):

| Step | isCloudManaged | onPremisesSyncEnabled |
| --- | --- | --- |
| Synced from AD DS | `false` | `true` |
| SOA converted to the cloud | `true` | `null` |
| Rolled back (until the sync client takes over) | `false` | `null` |
| Cloud-native group | `false` | `null` |

A filter on `onPremisesSyncEnabled eq true` alone hides every converted group. v1.08 therefore also queries `onPremisesSyncBehavior/isCloudManaged eq true`, which is the filter shown in [How to audit and monitor Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-source-of-authority-auditing-monitoring). If the second query fails in your tenant, the tool logs a warning and continues with synced groups only.

**Known gap**: a group that was rolled back but not yet picked up by the sync client (`isCloudManaged = false`, `onPremisesSyncEnabled = null`) looks like a cloud-native group. It appears again after the next sync cycle.

## Troubleshooting

- **Converted groups are missing**: check the log for "Could not retrieve converted groups". In Graph Explorer, test the query 2 URL above with the `ConsistencyLevel: eventual` header.
- **Conversion failures (403)**: make sure `Group-OnPremisesSyncBehavior.ReadWrite.All` is consented and that you have the Hybrid Administrator role. Reconnect to trigger the consent flow again.
- **Unverified results**: the PATCH reported an error and the state couldn't be read back. Check the group in the Entra admin center and run Search again.
- **Throttling**: the tool honours `Retry-After` and retries automatically. Progress may pause during throttling.
- **Module installation fails**: install manually with `Install-Module -Name Microsoft.Graph.Groups -Scope CurrentUser`. The tool also adds the CurrentUser module folder to `PSModulePath` when it is missing (for example on Exchange servers or with redirected Documents).

## Notes / Limitations

- Dynamic Distribution Groups and Microsoft 365 groups are not supported for SOA conversion.
- Groups are displayed in pages of 100.
- Optional: place a `logo.png` in this folder to display a logo in the header.

## Changelog

Versions 1.04 - 1.07 were contributed by a community contributor.

### Version 1.08
- **Converted groups in discovery**: discovery also returns groups whose SOA is already converted (`onPremisesSyncBehavior/isCloudManaged eq true`), so they are visible and can be rolled back after a new search.
- **CSV discovery fixes**:
  - Accepts converted groups.
  - Applies the search scope tick boxes and lists out-of-scope entries.
  - Replaces a broken fallback request with a correct direct retry.
- **Import CSV when groups are already loaded**: choose to load the CSV groups fresh from Graph or to filter the current list.
- **Disconnect** clears all loaded groups, nesting data and filters.
- **Select all**: unticked automatically when a row is clicked, and the confirmation states that all filtered groups (all pages) are processed.
- **Search box debounce**: the search applies 400 ms after typing stops, or immediately on Enter.
- **Throttling and cancel**: `Retry-After` is honoured for single requests and whole batch calls, and retries stop when a job is cancelled.
- **Buttons after a job**: Graph-dependent buttons are only re-enabled when still connected.
- **Confirmation dialog**: the message area is larger so long rollback warnings aren't cut off.
- **Clean-up**: removed unused functions, translated remaining comments to English, fixed the CSV import error message, and saved the script as UTF-8 with BOM.

### Version 1.07
- **Search scope**: three tick boxes (Security / Mail-enabled security / Distribution lists), replacing "Security groups only".
- **Synced groups only**: server-side filter plus a client-side safety net.
- **"Select all" tick box**: replaces "Apply to ALL filtered groups".
- **Removed** the "Hide Converted Groups" option.
- **CSV-driven discovery** when no groups are loaded yet.

### Version 1.06 (enterprise mode, 15K+ synced groups)
- **Discovery reworked**: only on-premises synced groups are retrieved (server-side filter).
- **JSON batching** with Retry-After aware throttling handling for SOA status and nesting lookups.
- **"Search Groups"** replaces automatic loading on connect and "Refresh Groups".
- **"Skip nesting analysis"** option (nesting is analyzed on demand).
- **Convert / rollback** in batches per nesting depth level, with Success / Failed / Unverified results.
- **Per-run CSV report**, progress with ETA and a "Cancel job" button. The window can't be closed while a job runs.
- **Fixed** empty group id/name in the SOA batch lookup.
- **ThreadPool** minimum threads raised at startup.

### Version 1.05
- **Folders**: logs are written to a `Logs` subfolder.
- **Filters**: group type, free-text search and CSV import.
- **Nested-children warning** checks all loaded groups.
- **"Exit" button**.
- **On-premises synced Security Groups** (not mail-enabled) are included.

### Version 1.04
- **Fixed** "no valid module file was found" after auto-install when the CurrentUser module folder is not in `PSModulePath`.

### Versions 1.0 - 1.03
See the [v1.03 README](../v1.03/README.md#changelog).

## Version

Current version: 1.08
