# Group SOA Conversion Tool

A PowerShell GUI tool for converting the *Group Source of Authority (SOA)* between on-premises Active Directory and Microsoft Entra ID (cloud) by toggling `isCloudManaged` via the Microsoft Graph `onPremisesSyncBehavior` API.

This is intended to support the approach described in:
- [Guidance for Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/concept-group-source-of-authority-guidance)
- [Configure Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure)

## Versions

The repository contains three versions, each in its own folder with its own README:

- **[v1.03](v1.03/README.md)**: the original tool.
- **[v1.04](v1.04/README.md)**: v1.03 plus a fix for the Microsoft Graph module not being found after a CurrentUser install (for example on Exchange servers or with redirected Documents). Otherwise it behaves exactly like v1.03.
- **[v1.08](v1.08/README.md)**: the enterprise version for large tenants. It also contains the v1.04 module fix.

The table compares the original line (v1.03 / v1.04) with v1.08:

| | [v1.03](v1.03/README.md) / [v1.04](v1.04/README.md) | [v1.08](v1.08/README.md) |
| --- | --- | --- |
| Script | `v1.03\Group-SOA-Conversion-Tool_1.03.ps1` / `v1.04\Group-SOA-Conversion-Tool_1.04.ps1` | `v1.08\Group-SOA-Conversion-Tool_1.08.ps1` |
| Target | Small and medium tenants, Exchange-focused | Large tenants (15K+ synced groups) |
| Module path fix after install | v1.04 only | Yes |
| Groups listed | All mail-enabled security groups and distribution groups (synced, converted and cloud-only) | On-premises synced and already converted groups: Security Groups, Mail-Enabled Security Groups, Distribution Groups. Cloud-only groups are excluded |
| Loading | Automatically after Connect | **Search Groups** with search scope tick boxes |
| Graph calls | One request per group | JSON batching (20 per call), Retry-After aware throttling |
| Nesting analysis | Always on load | On load, or on demand for the selected groups |
| Filters | Hide converted groups | Group type, free-text search, CSV import |
| CSV waves | - | Load or filter groups from a CSV (ObjectId / Mail / DisplayName) |
| Bulk selection | Multi-select on the current page | Multi-select, or **Select all** across all pages |
| Convert / rollback | One group at a time, in depth order | Batches per nesting depth level |
| Results | Success / Failed (unverifiable counted as success) | Success / Failed / Unverified / Skipped, plus a per-run CSV report |
| Progress / cancel | Status text | Progress with ETA, **Cancel job** |
| Log location | Script folder | `Logs` subfolder |

### Which version should I use?

- **v1.04** if you want the simple, proven tool for a modest number of Exchange groups. It lists everything mail-enabled, including cloud-only groups. Prefer it over v1.03, which is kept for reference.
- **v1.08** if you have many synced groups, also need plain (non-mail-enabled) security groups, want to run CSV-driven conversion waves, or need a report per run.

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Graph PowerShell module `Microsoft.Graph.Groups` (both versions install it in *CurrentUser* scope if it's missing)
- Microsoft Entra Connect Sync **2.5.76.0** or later, or Microsoft Entra Cloud Sync **1.1.1370.0** or later

### Permissions / Roles

- **Hybrid Administrator** role, to read and write `onPremisesSyncBehavior`. Without it the API returns 403.
- Graph permissions `Group.ReadWrite.All` and `Group-OnPremisesSyncBehavior.ReadWrite.All`.
- To grant admin consent: **Cloud Application Administrator** or **Application Administrator**.

Both versions trigger the consent flow when you click **Connect to Graph**. To grant consent manually:
1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com).
2. Go to **Enterprise Applications** > **Microsoft Graph Command Line Tools** > **Permissions**.
3. Click **Grant admin consent for [tenant name]**.

## Quick Start

```powershell
# Original version (with the module path fix)
cd .\v1.04
.\Group-SOA-Conversion-Tool_1.04.ps1 [-TenantId "00000000-0000-0000-0000-000000000000"]

# Enterprise version
cd .\v1.08
.\Group-SOA-Conversion-Tool_1.08.ps1 [-TenantId "00000000-0000-0000-0000-000000000000"]
```

See the folder README for the full usage, output files and Graph calls of each version.

## SOA Attribute States

After an SOA change, Microsoft Entra ID sets these attributes on the group ([source](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure#status-of-attributes-after-you-convert-soa)):

| Step | isCloudManaged | onPremisesSyncEnabled |
| --- | --- | --- |
| Synced from AD DS | `false` | `true` |
| SOA converted to the cloud | `true` | `null` |
| Rolled back (until the sync client takes over) | `false` | `null` |
| Cloud-native group | `false` | `null` |

This is why v1.08 discovers groups with two queries: `onPremisesSyncEnabled eq true` for synced groups, and `onPremisesSyncBehavior/isCloudManaged eq true` for converted groups ([auditing and monitoring docs](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-source-of-authority-auditing-monitoring)).

## Repository Layout

```
README.md      this overview
LICENSE
v1.03/         original version (script, logo, README)
v1.04/         original version + module path fix (script, logo, README)
v1.08/         enterprise version (script, logo, README)
```

Logs, reports and CSV exports contain tenant data and are excluded by `.gitignore`.

## Credits

- Original tool and v1.0 - v1.04, v1.08: Hans Christian Andersen ([blog](http://www.hcandersen.net), [LinkedIn](https://www.linkedin.com/in/hanschrandersen/)).
- v1.04 - v1.07 enterprise rework (batching, CSV waves, reports, cancel): contributed by a community contributor.

## License

MIT License, see [LICENSE](LICENSE). This script is provided AS-IS, with no warranty - use at your own risk.
