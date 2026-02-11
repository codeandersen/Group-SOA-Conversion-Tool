# Group SOA Conversion Tool

A PowerShell GUI tool for converting the *Group Source of Authority (SOA)* between on-premises Active Directory and Microsoft Entra ID (cloud) by toggling `isCloudManaged` via the Microsoft Graph `onPremisesSyncBehavior` API.

This tool focuses specifically on **Exchange on-premises relevant groups**: Mail-Enabled Security Groups and Distribution Groups (not Dynamic Distribution Groups or Microsoft 365 Groups).

This is intended to support the approach described in:
- [Guidance for Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/concept-group-source-of-authority-guidance)
- [Configure Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure)

## Features

- **Modern GUI interface**: Clean Windows Forms interface with responsive design and modern styling
- **Automatic module installation**: Checks for `Microsoft.Graph.Groups` and installs it in *CurrentUser* scope if needed
- **Microsoft Graph connectivity**: Uses modern auth via `Connect-MgGraph`
- **Permission status indicator**: Real-time green/red indicator showing whether `Group-OnPremisesSyncBehavior.ReadWrite.All` is consented
- **Setup Graph Permissions button**: One-click consent flow for the required SOA permission
- **Exchange-focused group list**: Retrieves on-premises synced groups and displays only Mail-Enabled Security Groups and Distribution Groups
- **Nested group detection**: Automatically detects nested group relationships via Microsoft Graph and computes nesting depth
- **Smart conversion ordering**: When converting multiple groups, automatically sorts bottom-up (deepest nested first) to follow Microsoft's recommended approach
- **Nested group warnings**: Warns if you try to convert a parent group before its nested children
- **Smart rollback ordering**: When rolling back, processes top-down (parents first, then children)
- **Rollback safety reminder**: Reminds you to remove cloud users and access packages before rolling back
- **Batch conversion**: Multi-select groups and convert in bulk with confirmation dialogs
- **Cloud conversion**: Sets `isCloudManaged = true` via the `onPremisesSyncBehavior` API
- **On-prem rollback**: Sets `isCloudManaged = false` via the `onPremisesSyncBehavior` API
- **Pagination support**: Displays groups in pages of 100 with Previous/Next navigation
- **Logging + quick access**: Writes a timestamped log file and includes an **Open Log File** button
- **Connection management**: Connect, refresh, and disconnect from Microsoft Graph with status indicators
- **Logo support**: Displays custom logo (logo.png) if present in script directory
- **Responsive layout**: Automatically adjusts to window resizing

## Requirements

- Windows PowerShell 5.1 or later
- Microsoft Graph PowerShell module: `Microsoft.Graph.Groups`
- Connectivity to Microsoft Graph endpoints
- Appropriate Microsoft Entra ID permissions (see below)

### Prerequisites

Before using this tool, ensure:

1. **Microsoft Entra Connect Sync** version **2.5.76.0** or later is installed, OR
2. **Microsoft Entra Cloud Sync** version **1.1.1370.0** or later is installed

### Permissions / Roles

At minimum, the signed-in admin account must have:

- **Hybrid Identity Administrator** role (least privileged role required for reading and writing `onPremisesSyncBehavior` — without this role, the API returns 403 Forbidden)
- **Group.ReadWrite.All** Graph permission (for reading and writing group properties)
- **Group-OnPremisesSyncBehavior.ReadWrite.All** Graph permission (for toggling the SOA)

To grant admin consent for the Graph permissions, you need one of:
- **Cloud Application Administrator** (recommended, least-privilege for consent)
- **Application Administrator**

### Setting Up Graph Permissions

The tool has a built-in **"Setup Graph Permissions"** button that triggers the consent flow. However, if that doesn't work (e.g., you lack the Application Administrator role), you can grant consent manually:

#### Option 1: Microsoft Entra Admin Center
1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com)
2. Navigate to **Enterprise Applications**
3. Find **Microsoft Graph Command Line Tools**
4. Go to **Permissions**
5. Click **Grant admin consent for [tenant name]**
6. Accept the `Group-OnPremisesSyncBehavior.ReadWrite.All` permission

#### Option 2: Graph Explorer
1. Open [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
2. Sign in as an Application Administrator or Cloud Application Administrator
3. Select the profile icon → **Consent to permissions**
4. Search for `Group-OnPremisesSyncBehavior` and select **Consent**

## Usage

1. **Run the Script**:
   ```powershell
   .\Group-SOA-Conversion-Tool.ps1 -TenantId "00000000-0000-0000-0000-000000000000"
   ```
   The `-TenantId` parameter is **required** to ensure you always connect to the correct tenant. The tenant ID is displayed in the title bar once connected.

2. **Connect to Microsoft Graph**:
   - Click **"Connect to Graph"** button
   - Sign in with your admin credentials
   - The tool will automatically load Exchange-relevant on-premises synced groups
   - Button changes to "Connected" with green color upon successful connection

3. **Check Permission Status**:
   - Look at the permission indicator (top-right area):
     - **Green**: "Graph Permissions and Consent OK" — you're ready to convert
     - **Red**: "Graph Permissions and Consent Missing" — click "Setup Graph Permissions" first
   - Click **"Setup Graph Permissions"** if the indicator is red

4. **Review Groups**:
   - The grid shows: Display Name, Email, Group Type, Cloud Managed status, and Nesting Depth
   - **Nesting Depth 0** = leaf group (no nested children in scope) — convert these first
   - **Higher depth** = has nested children — convert after children are done
   - Use "Previous" and "Next" buttons to navigate through pages

5. **Convert Groups to Cloud Managed**:
   - Select one or multiple groups from the list (multi-select supported)
   - Click **"Convert to Cloud Managed"**
   - The tool will:
     - **Warn** if any selected parent group has unconverted nested children not in your selection
     - **Sort** the selected groups bottom-up (lowest nesting depth first)
     - Show the conversion order for your confirmation
   - Confirm the conversion when prompted
   - View batch conversion summary

6. **Roll Back Groups to On-Prem**:
   - Select one or multiple groups
   - Click **"Roll Back to On-Prem"**
   - The tool will sort groups top-down (highest nesting depth first) for rollback
   - **Important**: Remove cloud users from groups and remove groups from access packages before rolling back
   - Note: Rollback is only complete after the next Connect Sync cycle

7. **Refresh Group List**:
   - Click **"Refresh Groups"** to reload the group list and SOA status after conversions

8. **View Logs**:
   - Click **"Open Log File"** to view the session log in Notepad

9. **Disconnect**:
   - Click **"Disconnect from Graph"** when finished
   - Tool automatically disconnects when closing the window

## Understanding Nested Groups

The tool automatically detects nested group relationships. Here's how it works:

- If **GroupB** is a member of **GroupA**, then GroupB is a "child" and GroupA is a "parent"
- **GroupB** should be converted **before** GroupA
- The **Nesting Depth** column shows how deep in the hierarchy a group sits:
  - **Depth 0**: No nested children in scope (safe to convert first)
  - **Depth 1**: Has direct nested children
  - **Depth 2+**: Has nested children that themselves have nested children
- When converting, the tool automatically processes groups from **depth 0 upward**
- When rolling back, the tool automatically processes groups from **highest depth downward**

### Example

```
GroupA (Depth 2)
  └── GroupB (Depth 1)  ← member of GroupA
        └── GroupC (Depth 0)  ← member of GroupB
```

**Conversion order**: GroupC → GroupB → GroupA
**Rollback order**: GroupA → GroupB → GroupC

## Log Files

Log files are created in the same directory as the script with the naming format:
```
GroupSOAConversion_YYYYMMDD_HHMM.log
```

Logged operations include:
- Microsoft Graph module installation attempts
- Connection to Microsoft Graph
- Permission checks and consent flows
- Group retrieval and SOA status checks
- Nested group analysis
- Group conversions to cloud managed
- Group rollbacks to on-premises managed
- Any errors or warnings

## Graph API Calls

The tool executes the following Microsoft Graph API calls:

**Check SOA status**:
```
GET https://graph.microsoft.com/v1.0/groups/{ID}/onPremisesSyncBehavior?$select=isCloudManaged
```

**Convert to Cloud Managed**:
```
PATCH https://graph.microsoft.com/v1.0/groups/{ID}/onPremisesSyncBehavior
{ "isCloudManaged": true }
```

**Roll Back to On-Premises Managed**:
```
PATCH https://graph.microsoft.com/v1.0/groups/{ID}/onPremisesSyncBehavior
{ "isCloudManaged": false }
```

## How This Maps to the Microsoft Guidance

### What the tool does
1. Connects to Microsoft Graph and retrieves on-premises synced groups
2. Filters for Exchange-relevant groups (Mail-Enabled Security Groups and Distribution Groups)
3. Checks the current `isCloudManaged` status for each group
4. Detects nested group relationships and computes conversion order
5. Converts group SOA by calling the `onPremisesSyncBehavior` API

### What happens after conversion to cloud-managed?
Once a group is converted to cloud-managed (`isCloudManaged = true`):
- The group can be managed in Microsoft Entra ID
- On-premises changes to the group are **no longer synced** to the cloud
- The `blockOnPremisesSync` property is set to `true` on the Entra ID object
- Event ID 6956 appears in the Application log on the Connect Sync server, indicating the object isn't synced because its SOA is in the cloud

### Limitations (from Microsoft documentation)
- **No reconciliation**: Local AD DS changes to cloud-owned groups are overwritten when group provisioning runs
- **No dual write**: After converting a group's SOA, membership references aren't synced for nested groups
- **No SOA conversion of nested groups**: Only the specified group's SOA is converted; nested groups must be converted one-by-one (this tool handles this automatically)
- **No extension attributes (1-15)**: Not supported on cloud security groups after SOA conversion

## Troubleshooting

- **Module Installation Fails**: Ensure you have internet connectivity and appropriate permissions. You can manually install the module using:
  ```powershell
  Install-Module -Name Microsoft.Graph.Groups -Scope CurrentUser
  ```

- **Permission Status Shows Red**: Click "Setup Graph Permissions" or manually grant consent in the Entra admin center (see "Setting Up Graph Permissions" section above)

- **Connection Issues**: Verify your credentials have the necessary Microsoft Entra ID permissions

- **Conversion Failures (403 errors)**: Ensure `Group-OnPremisesSyncBehavior.ReadWrite.All` is consented — check the permission status indicator

- **Groups Not Showing**: The tool only shows on-premises synced, mail-enabled groups. Cloud-only groups, Microsoft 365 groups, and non-mail-enabled security groups are excluded by design.

## Notes / Limitations

- The tool intentionally filters for Exchange-relevant groups only: Mail-Enabled Security Groups and Distribution Groups synced from on-premises AD.
- Dynamic Distribution Groups are **not** supported for SOA conversion.
- Microsoft 365 (Unified) groups are excluded.
- Changes may take time to reflect depending on your environment and any directory sync / hybrid processes.
- Groups are displayed in pages of 100 for better performance with large group counts.
- The tool automatically disconnects from Microsoft Graph when the window is closed.
- Optional: Place a `logo.png` file in the same directory as the script to display a custom logo in the header.

## Reference

- [Guidance for Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/concept-group-source-of-authority-guidance)
- [Configure Group SOA](https://learn.microsoft.com/en-us/entra/identity/hybrid/how-to-group-source-of-authority-configure)
- [Get onPremisesSyncBehavior API](https://learn.microsoft.com/en-us/graph/api/onpremisessyncbehavior-get)
- [Update onPremisesSyncBehavior API](https://learn.microsoft.com/en-us/graph/api/onpremisessyncbehavior-update)

## Version

Current version: 1.0
